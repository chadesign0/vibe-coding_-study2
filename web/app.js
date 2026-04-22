/**
 * 병원별 배점표 뷰어 — data/scoring-data.json
 *
 * 병원 간 데이터는 섞이지 않음: JSON 월 항목은 (monthLabel + hospitalName) 슬롯별로 병합되며,
 * hospitalName 이 없는 항목은 포인트병원(레거시) 전용으로만 표시된다.
 */

const DEFAULT_API_BASE = "http://127.0.0.1:8080";

/** 주소창 ?api= 또는 ?apiBase=, 또는 localStorage scoringApiBase (배포용 공용 서버) */
function getConfiguredApiBase() {
  try {
    const p = new URLSearchParams(window.location.search);
    const q = (p.get("api") || p.get("apiBase") || "").trim();
    if (q) return q.replace(/\/+$/, "");
  } catch (_) {}
  try {
    const s = localStorage.getItem("scoringApiBase");
    if (s && String(s).trim()) return String(s).trim().replace(/\/+$/, "");
  } catch (_) {}
  return DEFAULT_API_BASE;
}

const SCORING_SERVER_UNAVAILABLE =
  "채점 서버에 연결하지 못했습니다.\n" +
  "① 배점표 프로젝트 폴더에서 터미널을 연 뒤 python server.py 실행\n" +
  "② 브라우저에서 http://127.0.0.1:8080 으로 이 페이지 열기\n" +
  "(팀 공용 서버가 있으면 주소에 ?api=https://서버주소 를 붙일 수 있습니다.)";

/** 같은 탭·다른 포트·file:// 등에서도 채점 서버로 이어지도록 후보 URL 목록 */
function isLocalDev() {
  const h = window.location.hostname;
  return h === "localhost" || h === "127.0.0.1" || h === "";
}

function getScoringDataUrls() {
  const base = getConfiguredApiBase();
  const list = [];
  if (window.location.protocol === "http:" || window.location.protocol === "https:") {
    list.push(new URL("/data/scoring-data.json", window.location.origin).href);
  }
  if (isLocalDev()) list.push(`${base}/data/scoring-data.json`);
  return [...new Set(list)];
}

function getEvidenceUrls() {
  const base = getConfiguredApiBase();
  const list = [];
  if (window.location.protocol === "http:" || window.location.protocol === "https:") {
    list.push(new URL("/data/last-run-evidence.json", window.location.origin).href);
  }
  if (isLocalDev()) list.push(`${base}/data/last-run-evidence.json`);
  return [...new Set(list)];
}

const TAB_BY_COL = {
  7: "powerlink",
  8: "bizsite",
  9: "map",
  10: "cafe",
  11: "blog",
  12: "news",
  13: "video",
  14: "web",
};

const HIDDEN_COLS = new Set([6]); // 연관 검색어 열 숨김

/** 네이버 통합검색 탭 채널 열 인덱스 (TAB_BY_COL 과 동일) */
const CAFE_SCORE_COL = 10;
const BLOG_SCORE_COL = 11;
/** '키워드별 합계' 열 (scripts/build_april_month.py HEADER 와 동일) */
const KEYWORD_TOTAL_COL = 15;
const SCORE_COL_MIN = 7;
const SCORE_COL_MAX = 14;

const TAB_LABEL = {
  powerlink: "파워링크(순위)",
  bizsite: "비즈사이트",
  map: "지도",
  cafe: "카페",
  blog: "블로그",
  news: "보도자료",
  video: "동영상",
  web: "웹(통합검색)",
};

/** 모든 병원 공통 월 탭(항상 표시). 데이터는 JSON에 해당 monthLabel·병원이 있을 때만 채워짐. */
const FIXED_MONTH_TABS = [
  "1월",
  "2월",
  "3월",
  "4월",
  "5월",
  "6월",
  "7월",
  "8월",
  "9월",
  "10월",
  "11월",
  "12월",
];

/** 새로고침 후에도 병원·탭·검색어 등 유지 */
const VIEW_STATE_KEY = "scoringViewerUi";
/** 최근 채점 소요시간(ms) 히스토리(예상 남은시간 계산용) */
const SCORE_TIME_HISTORY_KEY = "scoringUploadDurationsMs";
const MAX_ROWS_IN_ALL_SHEETS_VIEW = 2000;
const MAX_RENDER_ROWS = 1200;

const state = {
  data: null,
  evidence: {},
  evidenceRoot: null,
  /** API 로딩 전에도 병원 드롭다운이 비지 않게 기본값 유지(초기 클릭 레이스 방지) */
  hospitals: ["포인트병원"],
  availableHospitals: new Set(["포인트병원"]),
  hospitalIndex: 0,
  hospitalFilter: "",
  monthIndex: 0,
  sheetIndex: 0,
  filter: "",
  /** 표에서 카페·블로그 열 보기: 전체 | 카페만 | 블로그만 (채점·JSON 동일) */
  channelView: "all",
  /** 키워드 입력란에서 업로드 시 부여할 범위 (지역·전국·전체) */
  keywordUploadScope: "all",
};

let uploadInFlight = false;
let scoreStatusTimer = null;
let scoreStartedAt = 0;

function readViewState() {
  try {
    const raw = localStorage.getItem(VIEW_STATE_KEY);
    if (!raw) return null;
    const v = JSON.parse(raw);
    return v && typeof v === "object" ? v : null;
  } catch (_) {
    return null;
  }
}

function persistViewState() {
  try {
    localStorage.setItem(
      VIEW_STATE_KEY,
      JSON.stringify({
        hospitalName: currentHospitalName(),
        monthIndex: state.monthIndex,
        sheetIndex: state.sheetIndex,
        filter: state.filter,
        channelView: state.channelView,
        keywordUploadScope: state.keywordUploadScope,
        hospitalFilter: state.hospitalFilter,
      })
    );
  } catch (_) {}
}

/** 병원 목록이 준비된 뒤 호출: 저장된 병원·월·시트·필터 복원 */
function applySavedViewState() {
  const v = readViewState();
  if (!v) return;
  const name = String(v.hospitalName || "").trim();
  if (name) {
    const idx = state.hospitals.indexOf(name);
    if (idx >= 0) state.hospitalIndex = idx;
  }
  if (
    typeof v.monthIndex === "number" &&
    v.monthIndex >= 0 &&
    v.monthIndex < FIXED_MONTH_TABS.length
  ) {
    state.monthIndex = v.monthIndex;
  }
  if (typeof v.sheetIndex === "number" && v.sheetIndex >= -1) {
    state.sheetIndex = v.sheetIndex;
  }
  if (typeof v.filter === "string") {
    state.filter = v.filter;
    const inp = document.getElementById("filterInput");
    if (inp) inp.value = v.filter;
  }
  if (v.channelView === "all" || v.channelView === "cafe" || v.channelView === "blog") {
    state.channelView = v.channelView;
  }
  if (
    v.keywordUploadScope === "regional" ||
    v.keywordUploadScope === "national" ||
    v.keywordUploadScope === "other" ||
    v.keywordUploadScope === "all"
  ) {
    state.keywordUploadScope = v.keywordUploadScope;
  }
  if (typeof v.hospitalFilter === "string") {
    state.hospitalFilter = v.hospitalFilter;
    const hi = document.getElementById("hospitalFilterInput");
    if (hi) hi.value = v.hospitalFilter;
  }
}

function normalizeHeader(h) {
  if (!h || String(h).trim() === "") return "";
  return String(h).replace(/\s+/g, " ").trim();
}

/** JSON에 예전 헤더가 남아 있어도 파워링크 열 의미를 통일 */
function columnHeaderLabel(raw, colIdx) {
  if (colIdx === 0 && (!raw || String(raw).trim() === "")) return "번호";
  const n = normalizeHeader(raw);
  if (colIdx === 7 && n.includes("파워") && n.includes("링크")) return "파워링크(순위)";
  return raw == null || raw === "" ? "—" : normalizeHeader(raw);
}

function showToast(msg) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.hidden = false;
  clearTimeout(showToast._t);
  const ms = String(msg).length > 100 ? 9000 : 3200;
  showToast._t = setTimeout(() => {
    el.hidden = true;
  }, ms);
}

async function loadJson(url) {
  const res = await fetch(url, { cache: "no-store" });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

/** 손상·구형 JSON에서도 뷰어가 죽지 않도록 months 만 보정 */
function normalizeScoringData(raw) {
  if (!raw || typeof raw !== "object") return { months: [] };
  const m = raw.months;
  if (!Array.isArray(m)) return { ...raw, months: [] };
  return raw;
}

/** 여러 경로 순차 시도 (Live Server, file://, 다른 포트에서도 127.0.0.1:8080 데이터 사용) */
async function loadScoringDataFirstAvailable() {
  let lastErr = null;
  for (const url of getScoringDataUrls()) {
    try {
      return await loadJson(url);
    } catch (e) {
      lastErr = e;
    }
  }
  throw lastErr ?? new Error("데이터를 불러올 수 없습니다.");
}

async function loadEvidenceFirstAvailable() {
  for (const url of getEvidenceUrls()) {
    try {
      return await loadJson(url);
    } catch (_) {}
  }
  return { evidence: {} };
}

async function loadHospitals() {
  const base = getConfiguredApiBase();
  const urls = [];
  if (window.location.protocol === "http:" || window.location.protocol === "https:") {
    urls.push(new URL("/api/hospitals", window.location.origin).href);
  }
  if (isLocalDev()) urls.push(`${base}/api/hospitals`);
  for (const url of [...new Set(urls)]) {
    try {
      const data = await loadJson(url);
      const items = Array.isArray(data?.hospitals) ? data.hospitals : [];
      const available = Array.isArray(data?.availableHospitals) ? data.availableHospitals : ["포인트병원"];
      return {
        hospitals: items.length ? items : ["포인트병원"],
        availableHospitals: new Set(available.length ? available : ["포인트병원"]),
      };
    } catch (_) {}
  }
  return { hospitals: ["포인트병원"], availableHospitals: new Set(["포인트병원"]) };
}

function currentHospitalName() {
  return state.hospitals[state.hospitalIndex] || "포인트병원";
}

function canonicalHospitalName(name) {
  const n = String(name || "").trim();
  if (!n) return "포인트병원";
  if (n === "삼성본정형외과") return "삼성본병원";
  // 식별명(블로그목록)과 채점용 hospitalName 이 다를 때 (snu서울병원.txt)
  if (n === "SNU서울정형외과") return "SNU서울병원";
  return n;
}

function currentHospitalKey() {
  return canonicalHospitalName(currentHospitalName());
}

function isCurrentHospitalAvailable() {
  const name = currentHospitalName();
  const key = currentHospitalKey();
  return state.availableHospitals.has(name) || state.availableHospitals.has(key);
}

/** JSON 월 항목이 현재 선택 병원에 속하는지 (hospitalName 없음 = 포인트병원 레거시 데이터). */
function monthBelongsToHospital(m, hospitalName) {
  const target = canonicalHospitalName(hospitalName);
  const source = canonicalHospitalName(m?.hospitalName || "");
  const h = m?.hospitalName;
  if (h == null || h === "") return target === "포인트병원";
  return source === target;
}

/** 현재 병원 + 월 라벨에 해당하는 JSON 월 블록 (없으면 null). */
function monthRecordForHospitalMonth(monthLabel) {
  const all = state.data?.months ?? [];
  const name = currentHospitalName();
  const key = currentHospitalKey();
  // 포인트병원은 레거시(hospitalName 없음)와 신규(hospitalName=포인트병원)가 공존할 수 있어
  // 먼저 "명시 병원명" 항목을 우선 사용하고, 없을 때만 레거시를 사용한다.
  if (key === "포인트병원") {
    const explicit = all.find((m) => m.monthLabel === monthLabel && (m?.hospitalName || "") === "포인트병원");
    if (explicit) return explicit;
  }
  // 병원 별칭(예: 삼성본정형외과/삼성본병원)도 같은 슬롯으로 조회
  return all.find((m) => m.monthLabel === monthLabel && monthBelongsToHospital(m, key)) ?? null;
}

function currentMonthLabel() {
  return FIXED_MONTH_TABS[state.monthIndex] ?? FIXED_MONTH_TABS[0];
}

/** 오늘 날짜가 속한 달 탭 인덱스(1~4월만 사용, 5월 이후는 4월 탭). */
function monthTabIndexForToday() {
  const cur = new Date().getMonth() + 1;
  const clamped = Math.min(Math.max(cur, 1), FIXED_MONTH_TABS.length);
  return clamped - 1;
}

/** fetch 실패(연결 거절 등)는 ok:false로 돌려서 다음 URL 후보를 시도 */
async function tryPostSafe(url, formData) {
  try {
    const res = await fetch(url, { method: "POST", body: formData });
    const data = await res.json().catch(() => ({}));
    return { reached: true, res, data };
  } catch {
    return { reached: false, res: null, data: {} };
  }
}

/**
 * 채점 API POST. 현재 페이지 origin → 설정된 API 베이스 순으로 시도.
 * @param {string} path 예: "/api/upload-keywords"
 */
async function postForm(path, formData) {
  const base = getConfiguredApiBase();
  const urls = [];
  if (window.location.protocol === "http:" || window.location.protocol === "https:") {
    urls.push(new URL(path, window.location.origin).href);
  }
  urls.push(`${base}${path}`);
  const unique = [...new Set(urls)];

  let anyReached = false;
  let lastErrText = "";

  for (const url of unique) {
    const { reached, res, data } = await tryPostSafe(url, formData);
    if (!reached) continue;
    anyReached = true;

    if (res.ok) {
      if (data && data.ok === false) {
        throw new Error(data.message || "요청 처리에 실패했습니다.");
      }
      return data;
    }
    lastErrText = (data && data.message) || `HTTP ${res.status}`;
  }

  if (anyReached) {
    throw new Error(
      lastErrText ||
        "서버가 요청을 받았지만 처리하지 못했습니다. python server.py 실행 여부를 확인하세요."
    );
  }
  throw new Error(SCORING_SERVER_UNAVAILABLE);
}

async function postJsonForBlob(path, bodyObj) {
  const base = getConfiguredApiBase();
  const urls = [];
  if (window.location.protocol === "http:" || window.location.protocol === "https:") {
    urls.push(new URL(path, window.location.origin).href);
  }
  urls.push(`${base}${path}`);
  const unique = [...new Set(urls)];
  let anyReached = false;
  let lastErr = "요청 처리에 실패했습니다.";
  for (const url of unique) {
    try {
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(bodyObj || {}),
      });
      anyReached = true;
      if (!res.ok) {
        const msg = await res.text().catch(() => "");
        lastErr = msg || `HTTP ${res.status}`;
        continue;
      }
      return res.blob();
    } catch (_) {}
  }
  if (!anyReached) throw new Error(SCORING_SERVER_UNAVAILABLE);
  throw new Error(lastErr);
}

async function loadScoreTask(taskId) {
  const base = getConfiguredApiBase();
  const urls = [];
  if (window.location.protocol === "http:" || window.location.protocol === "https:") {
    urls.push(new URL(`/api/score-task/${encodeURIComponent(taskId)}`, window.location.origin).href);
  }
  urls.push(`${base}/api/score-task/${encodeURIComponent(taskId)}`);
  let lastErr = null;
  for (const url of [...new Set(urls)]) {
    try {
      return await loadJson(url);
    } catch (e) {
      lastErr = e;
    }
  }
  throw lastErr ?? new Error("작업 상태를 확인하지 못했습니다.");
}

async function waitForScoreTask(taskId, labelText) {
  const timeoutMs = 20 * 60 * 1000;
  const started = Date.now();
  let networkErrStreak = 0;
  while (true) {
    if (Date.now() - started > timeoutMs) {
      throw new Error(`${labelText || "채점"} 작업 시간이 너무 오래 걸립니다. 서버 상태를 확인해주세요.`);
    }
    let payload;
    try {
      payload = await loadScoreTask(taskId);
      networkErrStreak = 0;
    } catch (e) {
      networkErrStreak++;
      if (networkErrStreak >= 5) throw e;
      setUploadScoreStatus(`연결 재시도중… (${networkErrStreak}/5)`);
      await new Promise((r) => setTimeout(r, 3000));
      continue;
    }
    const task = payload?.task || {};
    const st = String(task.status || "");
    if (st === "queued") {
      setUploadScoreStatus("작업 대기중...");
    } else if (st === "running") {
      setUploadScoreStatus("채점중...");
    } else if (st === "succeeded") {
      return task;
    } else if (st === "failed") {
      const msg = String(task.message || task.error || "채점 실행 실패");
      throw new Error(msg);
    } else {
      setUploadScoreStatus("작업 상태 확인중...");
    }
    await new Promise((r) => setTimeout(r, 1200));
  }
}

function resolveEvidenceBlock(keyword) {
  const root = state.evidenceRoot || {};
  const name = currentHospitalKey();
  const fromH = root.byHospital?.[name]?.[keyword];
  if (fromH) return fromH;
  return root.evidence?.[keyword];
}

function scoreReasonText(keyword, colIdx, cellValue) {
  const tab = TAB_BY_COL[colIdx];
  if (!tab) return "";
  const label = TAB_LABEL[tab] || tab;
  const ev = resolveEvidenceBlock(keyword)?.[tab];
  const shown = Number(cellValue || 0);
  if (tab === "powerlink") {
    if (!ev) {
      return `${label}: 표시값 ${shown}\n근거 데이터가 아직 없습니다.`;
    }
    const source = ev.source === "api" ? "API" : ev.source === "manual" ? "수동입력" : "웹수집";
    const rank = Number(ev.matched_rank ?? ev.rank ?? 0);
    const pos = rank > 0 ? `${rank}위` : "미노출(0)";
    const base = `${label}: ${pos}\n${source} · 위에서부터 1위=첫 파워링크 광고`;
    const snippet = String(ev.matched_text || ev.top?.[0]?.text || "").replace(/\s+/g, " ").trim();
    if (!snippet) return base;
    return `${base}\n근거: ${snippet.slice(0, 90)}${snippet.length > 90 ? "..." : ""}`;
  }
  const score = shown;
  if (!ev) {
    return `${label} ${score}점\n근거 데이터가 아직 없습니다.`;
  }
  const source = ev.source === "api" ? "API" : ev.source === "manual" ? "수동입력" : "웹수집";
  const rank = Number(ev.matched_rank ?? ev.rank ?? 0);
  let rule = "";
  if (rank <= 0) rule = "미노출(0점)";
  else if (rank === 1) rule = "1위(3점)";
  else if (rank <= 5) rule = "2~5위(2점)";
  else if (rank <= 10) rule = "6~10위(1점)";
  else rule = "10위 밖(0점)";

  const base = `${label} ${score}점\n${source} 기준 ${rank > 0 ? `${rank}위` : "미노출"} · 규칙 ${rule}`;
  const snippet = String(ev.matched_text || ev.top?.[0]?.text || "").replace(/\s+/g, " ").trim();
  if (!snippet) return base;
  return `${base}\n근거: ${snippet.slice(0, 90)}${snippet.length > 90 ? "..." : ""}`;
}

function getCurrentMonthRecord() {
  return monthRecordForHospitalMonth(currentMonthLabel());
}

function getCurrentSheet() {
  const m = getCurrentMonthRecord();
  if (!m?.sheets?.length) return null;
  if (state.sheetIndex === -1) return null;
  const max = m.sheets.length - 1;
  if (state.sheetIndex > max) state.sheetIndex = max;
  return m.sheets[state.sheetIndex] ?? m.sheets[0];
}

function getCurrentSheetsForView() {
  const m = getCurrentMonthRecord();
  if (!m?.sheets?.length) return [];
  if (state.sheetIndex === -1) return m.sheets;
  const one = getCurrentSheet();
  return one ? [one] : [];
}

function channelHiddenScoreCol() {
  if (state.channelView === "all") return null;
  if (state.channelView === "cafe") return BLOG_SCORE_COL;
  if (state.channelView === "blog") return CAFE_SCORE_COL;
  return null;
}

function isTableColumnHidden(colIdx) {
  if (HIDDEN_COLS.has(colIdx)) return true;
  const hideCh = channelHiddenScoreCol();
  return hideCh !== null && colIdx === hideCh;
}

/** 빈 화면·안내 행용 가시 열 개수 */
function defaultVisibleColumnCount() {
  const maxIdx = 16;
  let n = 0;
  for (let i = 0; i < maxIdx; i++) {
    if (!isTableColumnHidden(i)) n++;
  }
  return Math.max(n, 1);
}

function rowKeywordChannel(month, keyword) {
  const m = month?.keywordChannels;
  if (!m || typeof m !== "object") return "all";
  const ch = m[String(keyword).trim()];
  if (ch === "cafe" || ch === "blog" || ch === "all") return ch;
  return "all";
}

/** 툴바 '전체'면 모든 행, '카페'·'블로그'면 해당 구분(및 전체·미지정)만 */
function rowMatchesChannelView(month, keyword) {
  if (state.channelView === "all") return true;
  const ch = rowKeywordChannel(month, keyword);
  if (ch === "all") return true;
  if (state.channelView === "cafe") return ch === "cafe";
  if (state.channelView === "blog") return ch === "blog";
  return true;
}

function filteredRows(sheet) {
  const month = getCurrentMonthRecord();
  const baseRows = (sheet.rows || []).filter((row) => {
    const keyword = String(row?.[2] ?? "").trim();
    return keyword !== "";
  });
  const byChannel = baseRows.filter((row) => {
    const keyword = String(row?.[2] ?? "").trim();
    return rowMatchesChannelView(month, keyword);
  });
  const q = state.filter.trim().toLowerCase();
  if (!q) return byChannel;
  return byChannel.filter((row) => {
    const region = String(row[1] ?? "").toLowerCase();
    const keyword = String(row[2] ?? "").toLowerCase();
    return region.includes(q) || keyword.includes(q);
  });
}

function dedupeRowsByKeyword(rows) {
  const seen = new Set();
  const out = [];
  for (const row of rows || []) {
    const kw = String(row?.[2] ?? "").trim();
    if (!kw || seen.has(kw)) continue;
    seen.add(kw);
    out.push(Array.isArray(row) ? [...row] : row);
  }
  out.forEach((row, idx) => {
    if (Array.isArray(row) && row.length > 0) row[0] = idx + 1;
  });
  return out;
}

function cellNumericScore(cell) {
  if (typeof cell === "number" && Number.isFinite(cell)) return cell;
  if (cell === "" || cell == null) return 0;
  const n = Number(cell);
  return Number.isFinite(n) ? n : 0;
}

/**
 * 현재 열 표시(전체/카페/블로그)에 맞는 행 점수 합.
 * — 전체: JSON의 키워드별 합계 열(15) 우선, 없으면 가시 채점 열만 합산.
 * — 카페·블로그만: 숨긴 채널 열을 제외한 채점 열만 합산.
 */
function rowScoreSumForCurrentView(row) {
  if (!Array.isArray(row)) return 0;
  if (state.channelView === "all") {
    const t = row[KEYWORD_TOTAL_COL];
    if (typeof t === "number" && Number.isFinite(t)) return t;
    if (t !== "" && t != null) {
      const n = Number(t);
      if (Number.isFinite(n)) return n;
    }
  }
  let s = 0;
  for (let c = SCORE_COL_MIN; c <= SCORE_COL_MAX; c++) {
    if (isTableColumnHidden(c)) continue;
    s += cellNumericScore(row[c]);
  }
  return s;
}

function sumRowsKeywordScores(rows) {
  let total = 0;
  for (const row of rows || []) {
    total += rowScoreSumForCurrentView(row);
  }
  return total;
}

/** 행 수 문구 + 키워드별 합계 배지(우측 패널 메타) */
function updatePanelStats({ rowLine, keywordTotal, showKeywordBadge }) {
  const rowCount = document.getElementById("rowCount");
  const valEl = document.getElementById("keywordTotalValue");
  const badge = document.getElementById("keywordTotalBadge");
  if (rowCount) rowCount.textContent = rowLine ?? "";
  if (valEl) {
    valEl.textContent = Number.isFinite(keywordTotal)
      ? keywordTotal.toLocaleString()
      : String(keywordTotal ?? "");
  }
  if (badge) badge.hidden = showKeywordBadge === false;
}

function renderTable() {
  const head = document.getElementById("scoreHead");
  const body = document.getElementById("scoreBody");
  const sourceInfo = document.getElementById("sourceInfo");
  const hospitalName = currentHospitalName();
  const monthLabel = currentMonthLabel();
  const month = getCurrentMonthRecord();
  const sheetsForView = getCurrentSheetsForView();
  const sheet = sheetsForView[0] || null;

  if (!isCurrentHospitalAvailable()) {
    head.innerHTML = "";
    body.innerHTML = "";
    sourceInfo.textContent = `${hospitalName} · ${monthLabel}`;
    updatePanelStats({ rowLine: "", keywordTotal: 0, showKeywordBadge: false });
    const colCount = defaultVisibleColumnCount();
    body.innerHTML = `<tr><td colspan="${colCount}" class="num">${escapeHtml(
      `${hospitalName} 데이터는 아직 준비 중입니다. (배점표가 병합된 병원만 조회 가능)`
    )}</td></tr>`;
    return;
  }

  if (!month || !month?.sheets?.length) {
    head.innerHTML = "";
    updatePanelStats({ rowLine: "행 0건", keywordTotal: 0, showKeywordBadge: true });
    sourceInfo.textContent = `${hospitalName} · ${monthLabel} · 이 달 배점표 없음`;
    const colCount = defaultVisibleColumnCount();
    body.innerHTML = `<tr><td colspan="${colCount}" class="num">${escapeHtml(
      `${hospitalName}의 ${monthLabel} 데이터가 아직 없습니다. 채점·병합 후 새로고침 하세요.`
    )}</td></tr>`;
    return;
  }

  const sheetLabel = state.sheetIndex === -1 ? "전체 시트" : (sheet?.title || "시트");
  sourceInfo.textContent = `${hospitalName} · ${month.monthLabel} · ${sheetLabel} · 원본 ${month.sourceFile}`;
  if (month?.quality?.level === "warn" || month?.quality?.level === "low") {
    const acc = Number(month?.quality?.accuracyPct);
    const warnTxt = Number.isFinite(acc) ? `⚠ 정확도 ${acc.toFixed(1)}%` : "⚠ 정확도 경고";
    sourceInfo.textContent = `${sourceInfo.textContent} · ${warnTxt}`;
  }
  const rawRows = sheetsForView.flatMap((s) => filteredRows(s));
  const dedupedRows = state.sheetIndex === -1 ? dedupeRowsByKeyword(rawRows) : rawRows;
  const keywordGrandTotal = sumRowsKeywordScores(dedupedRows);
  const allSheetsOverLimit = state.sheetIndex === -1 && dedupedRows.length > MAX_ROWS_IN_ALL_SHEETS_VIEW;
  const baseRows = allSheetsOverLimit ? dedupedRows.slice(0, MAX_ROWS_IN_ALL_SHEETS_VIEW) : dedupedRows;
  const overRenderLimit = baseRows.length > MAX_RENDER_ROWS;
  const rows = overRenderLimit ? baseRows.slice(0, MAX_RENDER_ROWS) : baseRows;
  updatePanelStats({
    rowLine: `행 ${rows.length.toLocaleString()}건`,
    keywordTotal: keywordGrandTotal,
    showKeywordBadge: true,
  });

  const headerRow = sheet?.header || [];
  const ths = headerRow
    .map((h, i) => {
      if (isTableColumnHidden(i)) return "";
      const label = columnHeaderLabel(h, i);
      return `<th scope="col">${escapeHtml(label)}</th>`;
    })
    .join("");

  head.innerHTML = `<tr>${ths}</tr>`;

  // 상대평가: 현재 행들의 합계(col 15) 기준 상위/중위/하위 33%
  const totals = rows.map((row) => (typeof row?.[KEYWORD_TOTAL_COL] === "number" ? row[KEYWORD_TOTAL_COL] : 0));
  const sorted = [...totals].sort((a, b) => a - b);
  const lo = sorted[Math.floor(sorted.length * 0.33)] ?? 0;
  const hi = sorted[Math.floor(sorted.length * 0.67)] ?? 0;
  const rowTierClass = (total) => {
    if (typeof total !== "number" || total === 0) return "";
    if (total >= hi) return "row-high";
    if (total >= lo) return "row-mid";
    return "row-low";
  };

  body.innerHTML = rows
    .map((row) => {
      const keyword = String(row?.[2] ?? "").trim();
      const r = Array.isArray(row) ? row : [];
      const rowTotal = r[KEYWORD_TOTAL_COL];
      const tierCls = rowTierClass(rowTotal);
      // 헤더 열 개수 기준(행이 짧으면 빈 칸·길면 잘림)
      const cells = headerRow
        .map((_, colIdx) => {
          const cell = r[colIdx];
          if (isTableColumnHidden(colIdx)) return "";
          const isScoreCol = colIdx >= 7 && colIdx <= 14;
          const isTotalCol = colIdx === KEYWORD_TOTAL_COL;
          const isNumCol = [0, 4, 5, 15].includes(colIdx);
          const isKeywordCol = colIdx === 2;
          let cls = "";
          if (isScoreCol) cls = "score";
          else if (isTotalCol) cls = "score total num";
          else if (isNumCol) cls = "num";

          const empty = cell === null || cell === undefined || cell === "";
          let inner;
          if (empty) {
            inner = "—";
            if (isScoreCol || isTotalCol) cls += " is-empty";
          } else if (isTotalCol && typeof cell === "number") {
            inner = `<span class="countup-num" data-target="${cell}">0</span>`;
          } else if (typeof cell === "number") {
            inner = String(cell);
          } else {
            const text = String(cell);
            if (isKeywordCol) {
              const href = `https://search.naver.com/search.naver?query=${encodeURIComponent(text)}`;
              inner = `<a href="${href}" target="_blank" rel="noopener noreferrer">${escapeHtml(text)}</a>`;
            } else {
              inner = escapeHtml(text);
            }
          }
          const reason =
            !overRenderLimit && isScoreCol && keyword ? scoreReasonText(keyword, colIdx, cell) : "";
          const titleAttr = reason ? ` title="${escapeHtml(reason)}"` : "";
          const keywordAttr = isKeywordCol && keyword ? ` data-keyword="${escapeHtml(keyword)}"` : "";
          return `<td class="${cls}"${titleAttr}${keywordAttr}>${inner}</td>`;
        })
        .join("");
      return `<tr${tierCls ? ` class="${tierCls}"` : ""}>${cells}</tr>`;
    })
    .join("");
  runCountUp();

  if (allSheetsOverLimit || overRenderLimit) {
    const notes = [];
    if (allSheetsOverLimit) {
      notes.push(`전체 시트 상위 ${MAX_ROWS_IN_ALL_SHEETS_VIEW.toLocaleString()}건으로 축약`);
    }
    if (overRenderLimit) {
      notes.push(`브라우저 안정화를 위해 상위 ${MAX_RENDER_ROWS.toLocaleString()}건만 렌더`);
    }
    body.innerHTML += `<tr><td colspan="${defaultVisibleColumnCount()}" class="num">${notes.join(" · ")}. 검색어/시트로 범위를 좁혀주세요.</td></tr>`;
    updatePanelStats({
      rowLine: `행 ${rows.length.toLocaleString()}건 (원본 ${dedupedRows.length.toLocaleString()}건)`,
      keywordTotal: keywordGrandTotal,
      showKeywordBadge: true,
    });
  }
}

async function deleteKeywordFromTable(keyword) {
  const kw = String(keyword || "").trim();
  if (!kw) return;
  if (uploadInFlight) {
    showToast("다른 작업이 진행 중입니다. 잠시 후 다시 시도하세요.");
    return;
  }
  const ok = window.confirm(`"${kw}" 키워드를 삭제할까요?\n선택한 키워드 데이터만 즉시 제거됩니다.`);
  if (!ok) return;

  const form = new FormData();
  form.append("keyword", kw);
  uploadInFlight = true;
  try {
    showToast("키워드 삭제 중...");
    await postForm("/api/delete-keyword", form);
    await reloadDataAndRender();
    showToast(`삭제 완료: ${kw}`);
  } catch (e) {
    showToast("삭제 실패: " + (e?.message ?? e));
  } finally {
    uploadInFlight = false;
  }
}

function escapeHtml(s) {
  const text = String(s ?? "");
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function runCountUp() {
  const els = document.querySelectorAll("#scoreBody .countup-num");
  els.forEach((el, i) => {
    const target = Number(el.dataset.target);
    if (target === 0) { el.textContent = "0"; return; }
    const duration = 320;
    const delay = i * 30;
    setTimeout(() => {
      const start = performance.now();
      function step(now) {
        const t = Math.min((now - start) / duration, 1);
        const ease = 1 - Math.pow(1 - t, 3);
        el.textContent = Math.round(ease * target);
        if (t < 1) requestAnimationFrame(step);
        else el.textContent = target;
      }
      requestAnimationFrame(step);
    }, delay);
  });
}

function updateTabIndicator(wrap, indicatorClass) {
  const active = wrap.querySelector("[aria-selected='true']");
  let indicator = wrap.querySelector(`.${indicatorClass}`);
  if (!indicator) {
    indicator = document.createElement("span");
    indicator.className = indicatorClass;
    wrap.appendChild(indicator);
  }
  if (!active) { indicator.style.width = "0"; return; }
  requestAnimationFrame(() => {
    const wrapRect = wrap.getBoundingClientRect();
    const btnRect = active.getBoundingClientRect();
    indicator.style.left = `${btnRect.left - wrapRect.left}px`;
    indicator.style.width = `${btnRect.width}px`;
  });
}

function renderMonthTabs() {
  const wrap = document.getElementById("monthTabs");
  wrap.innerHTML = FIXED_MONTH_TABS.map(
    (label, i) =>
      `<button type="button" class="tab" role="tab" aria-selected="${i === state.monthIndex}" data-month="${i}">${escapeHtml(
        label
      )}</button>`
  ).join("");

  wrap.querySelectorAll(".tab").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.monthIndex = Number(btn.dataset.month);
      state.sheetIndex = 0;
      renderMonthTabs();
      renderSheetTabs();
      renderTable();
      persistViewState();
    });
  });
  updateTabIndicator(wrap, "tab-indicator");
}

function openHospitalMenu() {
  const menu = document.getElementById("hospitalMenu");
  const trigger = document.getElementById("hospitalTrigger");
  if (!menu || !trigger) return;
  menu.hidden = false;
  trigger.setAttribute("aria-expanded", "true");
}

function renderHospitalTabs() {
  const select = document.getElementById("hospitalSelect");
  const trigger = document.getElementById("hospitalTrigger");
  const triggerText = document.getElementById("hospitalTriggerText");
  const menu = document.getElementById("hospitalMenu");
  if (!select || !trigger || !triggerText || !menu) return;
  const hospitals =
    state.hospitals && state.hospitals.length ? state.hospitals : ["포인트병원"];
  const visibleHospitals = hospitals.map((name, i) => ({ name, i }));
  triggerText.textContent = currentHospitalName();
  menu.innerHTML = visibleHospitals
    .map(
      ({ name, i }) =>
        `<button type="button" class="hospital-select__option" role="option" aria-selected="${
          i === state.hospitalIndex
        }" data-hospital="${i}">${escapeHtml(
          name
        )}</button>`
    )
    .join("");
  if (!menu.innerHTML) {
    menu.innerHTML = `<div class="hospital-select__option hospital-select__empty" aria-selected="false">검색 결과가 없습니다.</div>`;
  }

  const closeMenu = () => {
    menu.hidden = true;
    trigger.setAttribute("aria-expanded", "false");
  };
  const openMenu = () => {
    menu.hidden = false;
    trigger.setAttribute("aria-expanded", "true");
  };

  trigger.onclick = () => {
    if (menu.hidden) openMenu();
    else closeMenu();
  };

  menu.querySelectorAll(".hospital-select__option").forEach((btn) => {
    if (btn.dataset.hospital === undefined) return;
    btn.addEventListener("click", () => {
      state.hospitalIndex = Number(btn.dataset.hospital);
      state.monthIndex = monthTabIndexForToday();
      state.sheetIndex = 0;
      renderHospitalTabs();
      renderMonthTabs();
      renderSheetTabs();
      renderTable();
      closeMenu();
      persistViewState();
    });
  });

  if (!renderHospitalTabs._boundOutsideClick) {
    document.addEventListener("click", (e) => {
      if (!select.contains(e.target)) closeMenu();
    });
    renderHospitalTabs._boundOutsideClick = true;
  }
}

function renderSheetTabs() {
  const quickWrap = document.getElementById("sheetQuickTabs");
  const m = getCurrentMonthRecord();
  const y = new Date().getFullYear();
  const fallbackSheets = [
    { title: `${y} 지역 PC` },
    { title: `${y} 지역 MOB` },
    { title: `${y} 전국 PC` },
    { title: `${y} 전국 MOB` },
    { title: `${y} 기타 PC` },
    { title: `${y} 기타 MOB` },
  ];
  const sheets = m?.sheets?.length ? m.sheets : fallbackSheets;
  if (state.sheetIndex >= sheets.length) state.sheetIndex = 0;
  const html = [
    `<button type="button" class="channel-tab" role="tab" aria-selected="${state.sheetIndex === -1}" data-sheet="-1">전체</button>`,
    ...sheets.map((s, i) => {
      const short = shortSheetLabel(s.title);
      return `<button type="button" class="channel-tab" role="tab" aria-selected="${i === state.sheetIndex}" data-sheet="${i}">${escapeHtml(short)}</button>`;
    }),
  ].join("");
  if (quickWrap) quickWrap.innerHTML = html;

  const bindSheetClick = (container, selector) => {
    if (!container) return;
    container.querySelectorAll(selector).forEach((btn) => {
      btn.addEventListener("click", () => {
        state.sheetIndex = Number(btn.dataset.sheet);
        renderSheetTabs();
        renderTable();
        persistViewState();
      });
    });
  };
  bindSheetClick(quickWrap, ".channel-tab");
  if (quickWrap) updateTabIndicator(quickWrap, "channel-tab-indicator");
}

function shortSheetLabel(title) {
  const t = title.replace(/\s+/g, " ").trim();
  if (/(지역|전국|기타).*(PC|MOB|모바일)/i.test(t)) {
    let area = "지역";
    if (t.includes("전국")) area = "전국";
    else if (t.includes("기타")) area = "기타";
    const dev =
      t.toUpperCase().includes("MOB") || t.includes("모바일") ? "MOB" : "PC";
    return `${area} ${dev}`;
  }
  return t;
}

/** 병원명 입력 + [검색] — 첫 일치 병원을 선택(드롭다운 목록은 항상 전체 표시) */
function applyHospitalNameFilter() {
  const hospitalInput = document.getElementById("hospitalFilterInput");
  if (!hospitalInput) return;
  const q = hospitalInput.value.trim();
  state.hospitalFilter = q;
  const hospitals =
    state.hospitals && state.hospitals.length ? state.hospitals : ["포인트병원"];
  const qLower = q.toLowerCase();
  const visible = hospitals
    .map((name, i) => ({ name, i }))
    .filter((x) => !qLower || x.name.toLowerCase().includes(qLower));
  if (visible.length && qLower) {
    state.hospitalIndex = visible[0].i;
    state.monthIndex = monthTabIndexForToday();
    state.sheetIndex = 0;
    renderMonthTabs();
    renderSheetTabs();
    renderTable();
  }
  renderHospitalTabs();
  openHospitalMenu();
  persistViewState();
}

function renderChannelTabs() {
  const wrap = document.getElementById("channelTabs");
  if (!wrap) return;
  const channels = [
    { id: "all", label: "전체" },
    { id: "cafe", label: "카페" },
    { id: "blog", label: "블로그" },
  ];
  wrap.innerHTML = channels
    .map(
      (c) =>
        `<button type="button" class="channel-tab" role="tab" aria-selected="${state.channelView === c.id}" data-channel="${c.id}">${escapeHtml(
          c.label
        )}</button>`
    )
    .join("");
  wrap.querySelectorAll(".channel-tab").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.channelView = btn.dataset.channel;
      renderChannelTabs();
      renderTable();
      persistViewState();
    });
  });
  updateTabIndicator(wrap, "channel-tab-indicator");
}

function renderUploadScopeTabs() {
  const wrap = document.getElementById("uploadScopeTabs");
  if (!wrap) return;
  const scopes = [
    { id: "all", label: "전체" },
    { id: "regional", label: "지역" },
    { id: "national", label: "전국" },
    { id: "other", label: "기타" },
  ];
  if (!["all", "regional", "national", "other"].includes(state.keywordUploadScope)) {
    state.keywordUploadScope = "all";
  }
  wrap.innerHTML = scopes
    .map(
      (s) =>
        `<button type="button" class="channel-tab" role="tab" aria-selected="${state.keywordUploadScope === s.id}" data-upload-scope="${s.id}">${escapeHtml(
          s.label
        )}</button>`
    )
    .join("");
  wrap.querySelectorAll("[data-upload-scope]").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.keywordUploadScope = btn.dataset.uploadScope;
      renderUploadScopeTabs();
      persistViewState();
    });
  });
}

function bindFilter() {
  const input = document.getElementById("filterInput");
  const hospitalInput = document.getElementById("hospitalFilterInput");
  const hospitalBtn = document.getElementById("hospitalFilterBtn");
  let t;
  input.addEventListener("input", () => {
    clearTimeout(t);
    t = setTimeout(() => {
      state.filter = input.value;
      renderTable();
      persistViewState();
    }, 120);
  });

  hospitalBtn?.addEventListener("click", () => applyHospitalNameFilter());
  hospitalInput?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      applyHospitalNameFilter();
    }
  });
}

function bindKeywordContextDelete() {
  const body = document.getElementById("scoreBody");
  if (!body) return;
  body.addEventListener("contextmenu", (e) => {
    const td = e.target.closest("td[data-keyword]");
    if (!td || !body.contains(td)) return;
    e.preventDefault();
    deleteKeywordFromTable(td.dataset.keyword || "");
  });
}

async function reloadDataAndRender() {
  const hs = await loadHospitals();
  state.hospitals = hs.hospitals;
  state.availableHospitals = hs.availableHospitals;
  state.data = normalizeScoringData(await loadScoringDataFirstAvailable());
  const ev = await loadEvidenceFirstAvailable();
  state.evidenceRoot = ev && typeof ev === "object" ? ev : {};
  state.evidence = state.evidenceRoot.evidence || {};
  if (state.hospitalIndex >= state.hospitals.length) state.hospitalIndex = 0;
  renderHospitalTabs();
  renderMonthTabs();
  renderSheetTabs();
  renderChannelTabs();
  renderTable();
  persistViewState();
}

function setUploadScoreStatus(text) {
  const el = document.getElementById("uploadScoreStatus");
  if (el) el.textContent = text || "";
}

function readRecentScoreDurations() {
  try {
    const raw = localStorage.getItem(SCORE_TIME_HISTORY_KEY);
    if (!raw) return [];
    const arr = JSON.parse(raw);
    if (!Array.isArray(arr)) return [];
    return arr.filter((x) => Number.isFinite(x) && x > 0).map((x) => Number(x));
  } catch (_) {
    return [];
  }
}

function saveRecentScoreDuration(ms) {
  if (!Number.isFinite(ms) || ms <= 0) return;
  const arr = readRecentScoreDurations();
  arr.push(ms);
  // 최근 6개만 유지
  const next = arr.slice(-6);
  try {
    localStorage.setItem(SCORE_TIME_HISTORY_KEY, JSON.stringify(next));
  } catch (_) {}
}

function formatMmSs(totalSec) {
  const sec = Math.max(0, Math.floor(totalSec));
  const m = Math.floor(sec / 60);
  const s = sec % 60;
  return `${m}:${String(s).padStart(2, "0")}`;
}

function estimatedScoreDurationMs() {
  const arr = readRecentScoreDurations();
  if (!arr.length) return 120000; // 기본 2분
  const avg = arr.reduce((a, b) => a + b, 0) / arr.length;
  return Math.max(30000, Math.min(900000, avg)); // 30초~15분 범위
}

function startScoreStatusTimer() {
  if (scoreStatusTimer) clearInterval(scoreStatusTimer);
  scoreStartedAt = Date.now();
  const estimatedMs = estimatedScoreDurationMs();
  const tick = () => {
    const elapsed = Date.now() - scoreStartedAt;
    const remainMs = estimatedMs - elapsed;
    if (remainMs > 0) {
      setUploadScoreStatus(`채점중 (예상 남은시간 ${formatMmSs(remainMs / 1000)})`);
    } else {
      setUploadScoreStatus("채점중 (곧 완료)");
    }
  };
  tick();
  scoreStatusTimer = setInterval(tick, 1000);
}

function stopScoreStatusTimer() {
  if (scoreStatusTimer) clearInterval(scoreStatusTimer);
  scoreStatusTimer = null;
}

/**
 * 키워드 채점 요청 후 표 갱신 (텍스트만 전송).
 * @param {"text"|"both"} mode
 */
async function runKeywordUpload(mode) {
  const textEl = document.getElementById("keywordsText");
  const uploadBtn = document.getElementById("uploadRunBtn");
  const text = (textEl?.value || "").trim();

  const form = new FormData();
  form.append("hospital_name", currentHospitalKey());
  form.append("month_label", currentMonthLabel());
  if (!text) {
    showToast(
      mode === "text" ? "키워드를 입력해주세요." : "키워드를 입력하세요."
    );
    return;
  }
  form.append("keywords_text", text);
  form.append("keyword_channel", "all");
  form.append("keyword_scope", state.keywordUploadScope || "all");

  if (uploadInFlight) return;
  uploadInFlight = true;
  if (uploadBtn) uploadBtn.disabled = true;
  startScoreStatusTimer();
  try {
    showToast("채점 중… 잠시만 기다려주세요.");
    const res = await postForm("/api/upload-keywords", form);
    let finishedTask = null;
    if (res?.accepted && res?.taskId) {
      finishedTask = await waitForScoreTask(res.taskId, "키워드 채점");
    }
    await reloadDataAndRender();
    saveRecentScoreDuration(Date.now() - scoreStartedAt);
    if (textEl) textEl.value = "";
    const monthTxt = res.monthLabel ? `${res.monthLabel} ` : "";
    const added = Number(res?.addedCount ?? 0);
    const q = finishedTask?.quality;
    if ((q?.level === "warn" || q?.level === "low") && Number.isFinite(Number(q?.accuracyPct))) {
      showToast(`정확도 경고: ${Number(q.accuracyPct).toFixed(1)}% (반영은 완료됨)`);
    }
    if (added <= 0) {
      showToast(`${monthTxt}중복 키워드로 추가 0건 (총 ${res.count}건 유지)`);
    } else {
      showToast(`${monthTxt}${added}개 추가 완료 (총 ${res.count}건)`);
    }
  } catch (e) {
    showToast("채점 실패: " + (e?.message ?? e));
  } finally {
    stopScoreStatusTimer();
    setUploadScoreStatus("채점 끝");
    uploadInFlight = false;
    if (uploadBtn) uploadBtn.disabled = false;
  }
}

function bindUploadActions() {
  const uploadBtn = document.getElementById("uploadRunBtn");
  const rerunBtn = document.getElementById("uploadRerunBtn");
  const exportBtn = document.getElementById("exportExcelBtn");
  const textEl = document.getElementById("keywordsText");

  uploadBtn?.addEventListener("click", () => {
    runKeywordUpload("both");
  });

  rerunBtn?.addEventListener("click", async () => {
    if (uploadInFlight) return;
    uploadInFlight = true;
    uploadBtn && (uploadBtn.disabled = true);
    rerunBtn.disabled = true;
    startScoreStatusTimer();
    try {
      showToast("전체 재채점 중… 잠시만 기다려주세요.");
      const form = new FormData();
      form.append("hospital_name", currentHospitalKey());
      form.append("month_label", currentMonthLabel());
      const res = await postForm("/api/run-score", form);
      if (res?.accepted && res?.taskId) {
        await waitForScoreTask(res.taskId, "재채점");
      } else if (res && res.ok === false) {
        throw new Error(res.message || "재채점 실패");
      }
      await reloadDataAndRender();
      saveRecentScoreDuration(Date.now() - scoreStartedAt);
      showToast("전체 재채점 완료");
    } catch (e) {
      showToast("재채점 실패: " + (e?.message ?? e));
    } finally {
      stopScoreStatusTimer();
      setUploadScoreStatus("채점 끝");
      uploadInFlight = false;
      uploadBtn && (uploadBtn.disabled = false);
      rerunBtn.disabled = false;
    }
  });

  exportBtn?.addEventListener("click", async () => {
    const headCells = Array.from(document.querySelectorAll("#scoreHead th"));
    const bodyRows = Array.from(document.querySelectorAll("#scoreBody tr"));
    if (!headCells.length || !bodyRows.length) {
      showToast("내보낼 표 데이터가 없습니다.");
      return;
    }
    const headers = headCells.map((th) => (th.textContent || "").trim());
    const rows = bodyRows
      .map((tr) =>
        Array.from(tr.querySelectorAll("td")).map((td) => (td.textContent || "").replace(/\s+/g, " ").trim())
      )
      .filter((r) => r.some((v) => v));
    if (!rows.length) {
      showToast("내보낼 표 데이터가 없습니다.");
      return;
    }
    const sheetLabel = state.sheetIndex === -1 ? "전체" : shortSheetLabel(getCurrentSheet()?.title || "시트");
    const fileName = `${currentHospitalName()}_${currentMonthLabel()}_${sheetLabel}_배점표.xlsx`;
    try {
      const blob = await postJsonForBlob("/api/export-xlsx", { headers, rows, fileName });
      const a = document.createElement("a");
      const url = URL.createObjectURL(blob);
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      showToast("엑셀 파일을 다운로드했습니다.");
    } catch (e) {
      showToast("엑셀 내보내기 실패: " + (e?.message ?? e));
    }
  });

  textEl?.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      runKeywordUpload("both");
    }
  });
}

/** 시안과 동일: 월·시트 가로 탭 영역 마우스 드래그 스크롤 */
function bindHorizontalDragScroll(el) {
  if (!el) return;
  let isDown = false;
  let startPageX = 0;
  let startScrollLeft = 0;
  el.style.cursor = "grab";
  el.addEventListener("mousedown", (e) => {
    if (e.button !== 0) return;
    isDown = true;
    el.style.cursor = "grabbing";
    startPageX = e.pageX;
    startScrollLeft = el.scrollLeft;
  });
  el.addEventListener("mouseleave", () => {
    isDown = false;
    el.style.cursor = "grab";
  });
  el.addEventListener("mouseup", () => {
    isDown = false;
    el.style.cursor = "grab";
  });
  el.addEventListener("mousemove", (e) => {
    if (!isDown) return;
    e.preventDefault();
    el.scrollLeft = startScrollLeft - (e.pageX - startPageX);
  });
}

async function init() {
  bindFilter();
  bindUploadActions();
  bindKeywordContextDelete();
  try {
    const hs = await loadHospitals();
    state.hospitals = hs.hospitals;
    state.availableHospitals = hs.availableHospitals;
    state.data = normalizeScoringData(await loadScoringDataFirstAvailable());
    const ev = await loadEvidenceFirstAvailable();
    state.evidenceRoot = ev && typeof ev === "object" ? ev : {};
    state.evidence = state.evidenceRoot.evidence || {};
    state.monthIndex = monthTabIndexForToday();
    state.sheetIndex = 0;
    applySavedViewState();
    if (state.hospitalIndex >= state.hospitals.length) state.hospitalIndex = 0;
  } catch (e) {
    showToast(
      "저장된 배점표를 불러오지 못했습니다. python server.py 실행 후 새로고침하거나, 아래에 키워드를 입력해 채점하세요."
    );
    state.data = { months: [] };
    state.hospitals = ["포인트병원"];
    state.availableHospitals = new Set(["포인트병원"]);
    state.evidenceRoot = null;
    state.evidence = {};
    state.monthIndex = monthTabIndexForToday();
    state.sheetIndex = 0;
    applySavedViewState();
    if (state.hospitalIndex >= state.hospitals.length) state.hospitalIndex = 0;
  }
  renderHospitalTabs();
  renderMonthTabs();
  renderSheetTabs();
  renderChannelTabs();
  renderUploadScopeTabs();
  renderTable();
  persistViewState();
  bindHorizontalDragScroll(document.getElementById("monthTabs"));
  bindHorizontalDragScroll(document.getElementById("sheetQuickTabs"));
}

init();
