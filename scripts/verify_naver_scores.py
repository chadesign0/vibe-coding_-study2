# -*- coding: utf-8 -*-
"""
클래스병원 3월 배점표 네이버 실시간 검증 스크립트
실행: python scripts/verify_naver_scores.py
결과: data/verify_result.json, data/verify_report.txt
"""
from __future__ import annotations
import json
import time
import random
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[1]
DATA_PATH = ROOT / "data" / "scoring-data.json"
RESULT_PATH = ROOT / "data" / "verify_result.json"
REPORT_PATH = ROOT / "data" / "verify_report.txt"

HOSPITAL_NAMES = ["클래스병원"]
HOSPITAL_DOMAINS = ["class2023.co.kr"]
BLOG_IDS = ["classlim2", "class231004"]

DELAY_MIN = 2.0
DELAY_MAX = 3.5


def contains_hospital(text: str) -> bool:
    t = (text or "").lower()
    return (
        any(n.lower() in t for n in HOSPITAL_NAMES) or
        any(d.lower() in t for d in HOSPITAL_DOMAINS)
    )


def contains_blog(text: str) -> bool:
    t = (text or "").lower()
    return any(b.lower() in t for b in BLOG_IDS)


def rank_to_score(rank: int) -> int:
    if not rank or rank <= 0: return 0
    if rank == 1: return 3
    if rank <= 5: return 2
    if rank <= 10: return 1
    return 0


def wait(page):
    time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))


def get_powerlink_rank(page) -> int:
    """파워링크 순위 (0=미노출). 의료심의필 포함 li 기준."""
    try:
        items = page.query_selector_all("li")
        pos = 0
        for li in items:
            try:
                text = li.inner_text() or ""
            except:
                continue
            if "의료심의필" in text:
                pos += 1
                if contains_hospital(text):
                    return pos
        return 0
    except:
        return 0


def get_ranked_items(page, check_fn, selector="li", max_pos=10) -> int:
    """일반 결과 목록에서 순위 반환 (0=미노출)."""
    try:
        items = page.query_selector_all(selector)
        pos = 0
        for el in items:
            try:
                text = el.inner_text() or ""
            except:
                continue
            text = text.strip()
            if len(text) < 5:
                continue
            pos += 1
            if check_fn(text):
                return pos
            if pos >= max_pos:
                break
        return 0
    except:
        return 0


def check_section(page, url: str, check_fn, max_pos=10) -> int:
    """URL 이동 후 순위 확인."""
    try:
        page.goto(url, timeout=20000)
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        wait(page)
        return get_ranked_items(page, check_fn, max_pos=max_pos)
    except Exception as e:
        return -1  # 오류


def check_map(page, keyword: str) -> int:
    """지도 - local 검색."""
    url = f"https://search.naver.com/search.naver?where=local&query={keyword}"
    try:
        page.goto(url, timeout=20000)
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        wait(page)
        return get_ranked_items(page, contains_hospital, max_pos=10)
    except:
        return -1


def check_blog(page, keyword: str) -> int:
    """블로그 - 클래스병원 블로그 ID로 확인."""
    url = f"https://search.naver.com/search.naver?where=blog&query={keyword}"
    try:
        page.goto(url, timeout=20000)
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        wait(page)
        return get_ranked_items(page, contains_blog, max_pos=10)
    except:
        return -1


def check_news(page, keyword: str) -> int:
    """보도자료(뉴스)."""
    url = f"https://search.naver.com/search.naver?where=news&query={keyword}"
    try:
        page.goto(url, timeout=20000)
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        wait(page)
        return get_ranked_items(page, contains_hospital, max_pos=10)
    except:
        return -1


def check_video(page, keyword: str) -> int:
    """동영상."""
    url = f"https://search.naver.com/search.naver?where=video&query={keyword}"
    try:
        page.goto(url, timeout=20000)
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        wait(page)
        return get_ranked_items(page, contains_hospital, max_pos=10)
    except:
        return -1


def check_web(page, keyword: str) -> int:
    """웹(통합검색)."""
    url = f"https://search.naver.com/search.naver?where=web&query={keyword}"
    try:
        page.goto(url, timeout=20000)
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        wait(page)
        return get_ranked_items(page, contains_hospital, max_pos=10)
    except:
        return -1


def check_bizsite(page, keyword: str) -> int:
    """비즈사이트 - 메인 페이지에서 확인."""
    try:
        # 메인 페이지는 이미 로드됐을 수 있으므로 URL 체크
        if keyword not in (page.url or ""):
            url = f"https://search.naver.com/search.naver?query={keyword}"
            page.goto(url, timeout=20000)
            page.wait_for_load_state("domcontentloaded", timeout=15000)
            wait(page)
        # 비즈사이트는 의료심의필 없이 class2023 도메인으로 확인
        items = page.query_selector_all("li")
        pos = 0
        for li in items:
            try:
                text = li.inner_text() or ""
                html = li.inner_html() or ""
            except:
                continue
            if "의료심의필" in text:
                continue  # 파워링크 제외
            if contains_hospital(text + html) and len(text) > 10:
                pos += 1
                return pos
        return 0
    except:
        return -1


def load_keywords() -> list[dict]:
    """scoring-data.json 에서 클래스병원 3월 지역PC 키워드+점수 추출."""
    data = json.loads(DATA_PATH.read_text(encoding="utf-8"))
    results = []
    seen = set()
    for m in data.get("months", []):
        if m.get("hospitalName") != "클래스병원" or m.get("monthLabel") != "3월":
            continue
        for sheet in m.get("sheets", []):
            for row in sheet.get("rows", []):
                if not row or len(row) < 15 or not str(row[2]).strip():
                    continue
                kw = str(row[2]).strip()
                sheet_key = sheet["key"]
                entry_key = f"{kw}|{sheet_key}"
                if entry_key in seen:
                    continue
                seen.add(entry_key)
                results.append({
                    "keyword": kw,
                    "region": str(row[1] or ""),
                    "sheet": sheet_key,
                    "stored": {
                        "powerlink": row[7] or 0,
                        "bizsite": row[8] or 0,
                        "map": row[9] or 0,
                        "cafe": row[10] or 0,
                        "blog": row[11] or 0,
                        "news": row[12] or 0,
                        "video": row[13] or 0,
                        "web": row[14] or 0,
                    }
                })
    return results


def main():
    keywords = load_keywords()
    total_kw = len(keywords)
    print(f"총 {total_kw}개 키워드 검증 시작 (예상 소요시간: {total_kw * 15 // 60}~{total_kw * 20 // 60}분)", flush=True)

    results = []
    match_count = 0
    mismatch_count = 0
    total_checks = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            locale="ko-KR",
            viewport={"width": 1280, "height": 900}
        )
        page = context.new_page()

        for idx, entry in enumerate(keywords):
            kw = entry["keyword"]
            stored = entry["stored"]
            print(f"[{idx+1}/{total_kw}] {kw} ({entry['sheet']})", flush=True)

            actual = {}
            mismatches = []

            # 1. 파워링크 + 비즈사이트 메인 페이지
            try:
                page.goto(f"https://search.naver.com/search.naver?query={kw}", timeout=20000)
                page.wait_for_load_state("domcontentloaded", timeout=15000)
                wait(page)
                pl_rank = get_powerlink_rank(page)
                actual["powerlink"] = pl_rank
            except:
                pl_rank = -1
                actual["powerlink"] = -1

            # 비즈사이트 (메인 페이지 재사용)
            actual["bizsite"] = check_bizsite(page, kw)

            # 2. 지도
            actual["map"] = check_map(page, kw)

            # 3. 블로그
            actual["blog"] = check_blog(page, kw)

            # 4. 보도자료
            actual["news"] = check_news(page, kw)

            # 5. 동영상
            actual["video"] = check_video(page, kw)

            # 6. 웹
            actual["web"] = check_web(page, kw)

            # 비교 (카페는 제외 - API 없이 정확한 순위 측정 어려움)
            for tab in ["bizsite", "map", "blog", "news", "video", "web"]:
                s_score = stored.get(tab, 0)
                a_rank = actual.get(tab, -1)
                if a_rank == -1:
                    continue  # 오류는 스킵
                a_score = rank_to_score(a_rank)
                total_checks += 1
                if s_score == a_score:
                    match_count += 1
                else:
                    mismatch_count += 1
                    mismatches.append({
                        "tab": tab,
                        "stored_score": s_score,
                        "actual_rank": a_rank,
                        "actual_score": a_score,
                    })

            # 파워링크는 순위 직접 비교
            s_pl = stored.get("powerlink", 0)
            a_pl = actual.get("powerlink", -1)
            if a_pl != -1:
                total_checks += 1
                if s_pl == a_pl:
                    match_count += 1
                else:
                    mismatch_count += 1
                    mismatches.append({
                        "tab": "powerlink",
                        "stored_rank": s_pl,
                        "actual_rank": a_pl,
                    })

            entry_result = {
                "keyword": kw,
                "region": entry["region"],
                "sheet": entry["sheet"],
                "stored": stored,
                "actual": actual,
                "mismatches": mismatches,
            }
            results.append(entry_result)

            if mismatches:
                for mm in mismatches:
                    print(f"  [불일치] {mm}", flush=True)
            else:
                print(f"  [일치]", flush=True)

            # 진행 저장 (중간 결과)
            if (idx + 1) % 10 == 0:
                RESULT_PATH.write_text(
                    json.dumps({"results": results, "progress": idx + 1, "total": total_kw}, ensure_ascii=False, indent=2),
                    encoding="utf-8"
                )

        browser.close()

    # 최종 저장
    accuracy = match_count / total_checks * 100 if total_checks else 0
    summary = {
        "total_keywords": total_kw,
        "total_checks": total_checks,
        "match": match_count,
        "mismatch": mismatch_count,
        "accuracy_pct": round(accuracy, 1),
        "results": results,
    }
    RESULT_PATH.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    # 리포트 작성
    lines = [
        "=== 클래스병원 3월 배점표 네이버 실시간 검증 리포트 ===",
        f"총 키워드: {total_kw}개",
        f"총 검증 항목: {total_checks}개",
        f"일치: {match_count}개",
        f"불일치: {mismatch_count}개",
        f"정확도: {accuracy:.1f}%",
        "",
        "=== 불일치 목록 ===",
    ]
    for r in results:
        for mm in r.get("mismatches", []):
            lines.append(f"[{r['sheet']}] {r['keyword']} / {mm.get('tab','?')}: 저장={mm.get('stored_score', mm.get('stored_rank','?'))} / 현재={mm.get('actual_score', mm.get('actual_rank','?'))} (순위:{mm.get('actual_rank','?')})")

    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")
    print("\n" + "\n".join(lines[:8]))
    print(f"\n상세 결과: {RESULT_PATH}")
    print(f"리포트: {REPORT_PATH}")


if __name__ == "__main__":
    main()
