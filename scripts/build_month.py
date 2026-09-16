# -*- coding: utf-8 -*-
"""병원별 배점표 생성: 네이버 API + 웹 파싱 자동 채점(근거 저장)."""
from __future__ import annotations

import base64
import hashlib
import hmac
import html
import json
import os
import random
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote_plus, urlparse

import requests
from bs4 import BeautifulSoup

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None

ROOT = Path(__file__).resolve().parents[1]
if load_dotenv:
    load_dotenv(ROOT / ".env")

HEADER = ["", "지역", "키워드", "계", "월간조회수(pc)", "월간조회수(모바일)", "연관   검색어", "파워링크(순위)", "비즈    사이트", "지도", "카페", "블로그", "보도    자료", "동영상", "웹", "키워드별 합계"]
SHEETS_META = [
    ("regional-pc", "2026 지역 PC"),
    ("regional-mob", "2026 지역 MOB"),
    ("national-pc", "2026 전국 PC"),
    ("national-mob", "2026 전국 MOB"),
    ("other-pc", "2026 기타 PC"),
    ("other-mob", "2026 기타 MOB"),
]
COL_BY_TAB = {"powerlink": 7, "bizsite": 8, "map": 9, "cafe": 10, "blog": 11, "news": 12, "video": 13, "web": 14}
POWERLINK_COL = COL_BY_TAB["powerlink"]
TAB_ENDPOINT = {"map": "local", "cafe": "cafearticle", "blog": "blog", "news": "news", "video": "video", "web": "webkr"}
SEARCH_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept-Encoding": "gzip, deflate, br",
    "Referer": "https://www.naver.com/",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}
MOBILE_SEARCH_HEADERS = {
    **SEARCH_HEADERS,
    "User-Agent": (
        "Mozilla/5.0 (Linux; Android 14; Pixel 8) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Mobile Safari/537.36"
    ),
}

# 스레드별 Session 재사용 (TCP 커넥션 풀 유지, 쿠키 자동 관리)
_thread_local = threading.local()


def _get_web_session(device: str = "pc") -> requests.Session:
    attr = "mobile_session" if device == "mobile" else "pc_session"
    if not getattr(_thread_local, attr, None):
        s = requests.Session()
        s.headers.update(MOBILE_SEARCH_HEADERS if device == "mobile" else SEARCH_HEADERS)
        setattr(_thread_local, attr, s)
    return getattr(_thread_local, attr)


TOTAL_COL = 15


def decode_response_text(r: requests.Response) -> str:
    """
    네이버 응답에서 한글이 깨지면 병원명 매칭이 실패할 수 있어 UTF-8 우선 디코딩한다.
    """
    try:
        return r.content.decode("utf-8", errors="ignore")
    except Exception:
        enc = (r.apparent_encoding or r.encoding or "utf-8").strip()
        try:
            return r.content.decode(enc, errors="ignore")
        except Exception:
            return r.text


def normalize_text(s: str) -> str:
    return re.sub(r"\s+", "", (s or "").lower())


def extract_domain_tokens(domain_or_url: str) -> list[str]:
    raw = (domain_or_url or "").strip().lower()
    if not raw:
        return []
    host = raw
    if "://" in raw:
        host = (urlparse(raw).netloc or "").lower()
    else:
        host = raw.split("/")[0].lower()
    host = host.split(":")[0].strip(".")
    if not host:
        return []
    vals = {host}
    if host.startswith("www."):
        vals.add(host[4:])
    return [v for v in vals if v]


def build_match_tokens(names: list[str], domains: list[str]) -> list[str]:
    tokens = {normalize_text(n) for n in (names or []) if (n or "").strip()}
    for d in domains or []:
        for t in extract_domain_tokens(d):
            tokens.add(normalize_text(t))
    return [t for t in tokens if t]


def parse_blog_postdate(raw: Any) -> date | None:
    """네이버 블로그 검색 API `postdate`(보통 yyyymmdd 문자열)를 date로 변환."""
    s = str(raw or "").strip()
    if len(s) >= 8 and s[:8].isdigit():
        try:
            return date(int(s[:4]), int(s[4:6]), int(s[6:8]))
        except ValueError:
            return None
    return None


def parse_cafe_date(raw: Any) -> date | None:
    """네이버 카페 검색 API `date` 필드를 date로 변환. ISO / yyyymmdd 등 다양한 형식 지원."""
    s = str(raw or "").strip()
    if not s:
        return None
    # yyyymmdd or yyyymmddHHmmss
    if s[:8].replace("-", "").isdigit():
        clean = s.replace("-", "").replace("T", "").replace(":", "")
        if len(clean) >= 8:
            try:
                return date(int(clean[:4]), int(clean[4:6]), int(clean[6:8]))
            except ValueError:
                pass
    # ISO: 2026-04-01T... or 2026-04-01 ...
    try:
        return date.fromisoformat(s[:10])
    except (ValueError, IndexError):
        pass
    return None


def cafe_item_in_scoring_month(item: dict[str, Any], year: int, month: int) -> bool:
    pd = parse_cafe_date(item.get("date"))
    if pd is None:
        return False
    return pd.year == year and pd.month == month


def parse_pubdate(raw: Any) -> date | None:
    """뉴스 API pubDate 필드 파싱. RFC 2822 형식: 'Mon, 01 Apr 2026 12:00:00 +0900'."""
    import email.utils
    s = str(raw or "").strip()
    if not s:
        return None
    try:
        parsed = email.utils.parsedate(s)
        if parsed:
            return date(parsed[0], parsed[1], parsed[2])
    except Exception:
        pass
    return parse_cafe_date(s)


def content_item_in_scoring_month(tab: str, item: dict[str, Any], year: int, month: int) -> bool:
    """news/video 탭 아이템의 게시 날짜가 채점 월인지 확인."""
    if tab == "news":
        pd = parse_pubdate(item.get("pubDate"))
    elif tab == "video":
        pd = parse_cafe_date(item.get("date"))
    else:
        return True
    if pd is None:
        return False
    return pd.year == year and pd.month == month


def blog_evidence_period(cfg: dict[str, Any]) -> tuple[int, int] | None:
    """
    블로그 채점 시 인정할 게시 연·월.
    날짜 필터 비활성화 — 병원 정보 노출 여부만으로 채점.
    """
    return None


def blog_item_in_scoring_month(item: dict[str, Any], year: int, month: int) -> bool:
    pd = parse_blog_postdate(item.get("postdate"))
    if pd is None:
        return False
    return pd.year == year and pd.month == month


def blog_author_blog_text_for_match(item: dict[str, Any]) -> str:
    """
    블로그 채점 매칭 전용: 작성자명·블로그 홈(bloggerlink)·글 URL(link)에서만 문자열을 만든다.
    제목·description은 제외(제목/본문만의 언급으로는 채점하지 않음).
    """
    parts: list[str] = []
    parts.append(strip_html(item.get("bloggername", "")))
    parts.append(strip_html(item.get("bloggerlink", "")))
    link = (item.get("link") or "").strip()
    if link:
        parts.append(strip_html(link))
        try:
            u = urlparse(link)
            netloc = (u.netloc or "").lower()
            if netloc:
                parts.append(netloc)
            path = (u.path or "").strip("/")
            if path:
                parts.extend(path.split("/")[:3])
        except Exception:
            pass
    return " ".join(p for p in parts if p)


def tokens_match_in_normalized(blob: str, match_tokens: list[str]) -> bool:
    n = normalize_text(blob)
    return any(t in n for t in match_tokens)


def naver_blog_id_from_url(url: str) -> str | None:
    """blog.naver.com/{id}/... 또는 m.blog.naver.com/{id}/... 의 블로그 아이디."""
    s = (url or "").strip()
    if not s:
        return None
    try:
        u = urlparse(s)
    except Exception:
        return None
    host = (u.netloc or "").lower()
    if "blog.naver.com" not in host:
        return None
    parts = [p for p in (u.path or "").split("/") if p]
    if not parts:
        return None
    first = parts[0].lower()
    if first in ("post.naver.com", "redirect.naver", "gate.naver.com", "naver.com"):
        return None
    return first


def text_has_naver_blog_id(text: str, bid: str) -> bool:
    """text 안에 blog.naver.com/{bid} 가 아이디 경계까지 정확히 등장하는지.

    단순 부분 문자열 비교는 'suca' 가 'sucaXXX' 블로그에도 매칭되므로,
    아이디 뒤에 아이디 문자(영문·숫자·_·-)가 이어지면 다른 블로그로 본다.
    """
    if not bid:
        return False
    pattern = r"blog\.naver\.com/" + re.escape(bid) + r"(?![A-Za-z0-9_-])"
    return re.search(pattern, text or "", re.IGNORECASE) is not None


def official_naver_blog_ids_from_config(cfg: dict[str, Any]) -> frozenset[str]:
    out: set[str] = set()
    for u in cfg.get("hospitalBlogBases") or []:
        bid = naver_blog_id_from_url(str(u).strip())
        if bid:
            out.add(bid)
    return frozenset(out)


def blog_item_matches_official_naver_blog(item: dict[str, Any], official_ids: frozenset[str]) -> bool:
    if not official_ids:
        return False
    for raw in (item.get("bloggerlink"), item.get("link")):
        bid = naver_blog_id_from_url(str(raw or ""))
        if bid and bid in official_ids:
            return True
    return False


def naver_cafe_id_from_url(url: str) -> str | None:
    """cafe.naver.com/{cafeId}/... 의 카페 아이디."""
    s = (url or "").strip()
    if not s:
        return None
    try:
        u = urlparse(s)
    except Exception:
        return None
    host = (u.netloc or "").lower()
    if "cafe.naver.com" not in host:
        return None
    parts = [p for p in (u.path or "").split("/") if p]
    if not parts:
        return None
    return parts[0].lower()


def official_naver_cafe_ids_from_config(cfg: dict[str, Any]) -> frozenset[str]:
    out: set[str] = set()
    for u in cfg.get("hospitalCafeBases") or []:
        cid = naver_cafe_id_from_url(str(u).strip())
        if cid:
            out.add(cid)
    return frozenset(out)


def cafe_item_matches_official_naver_cafe(item: dict[str, Any], official_ids: frozenset[str]) -> bool:
    if not official_ids:
        return False
    for raw in (item.get("link"), item.get("cafeurl"), item.get("cafename")):
        cid = naver_cafe_id_from_url(str(raw or ""))
        if cid and cid in official_ids:
            return True
        txt = normalize_text(strip_html(str(raw or "")))
        if any(oid in txt for oid in official_ids):
            return True
    return False


def strip_html(s: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html.unescape(s or ""))).strip()


def rank_to_score(rank: int | None) -> int | None:
    if rank is None:
        return None
    if rank < 1:
        return 0
    if rank <= 3:
        return 3
    if rank <= 5:   # 블로그/전용탭 상단 5위
        return 2
    if rank <= 10:  # 6~10위
        return 1
    return 0


def table_cell_for_tab(tab: str, raw_rank: int | None) -> int | None:
    """표에 넣을 값. 파워링크만 순위(0=미노출, 1~10), 웹은 노출=3/미노출=0, 나머지 탭은 점수(0~3)."""
    if tab == "powerlink":
        if raw_rank is None:
            return None
        return int(raw_rank)
    if tab == "web":
        if raw_rank is None:
            return None
        return 3 if raw_rank > 0 else 0
    return rank_to_score(raw_rank)


def parse_count(v: Any) -> int | None:
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    if s.startswith("<"):
        return 0
    d = re.sub(r"[^0-9]", "", s)
    return int(d) if d else None


def build_searchad_signature(ts: str, method: str, uri: str, secret: str) -> str:
    msg = f"{ts}.{method}.{uri}"
    dg = hmac.new(secret.encode("utf-8"), msg.encode("utf-8"), hashlib.sha256).digest()
    return base64.b64encode(dg).decode("utf-8")


def fetch_keyword_volumes_searchad(keywords: list[str], api_key: str, secret_key: str, customer_id: str) -> dict[str, dict[str, Any]]:
    endpoint = "/keywordstool"
    result: dict[str, dict[str, Any]] = {}
    for kw in keywords:
        # 검색광고 API 의 hintKeywords 는 띄어쓰기를 받지 않는다.
        # "발바닥 통증" 을 그대로 보내면 4xx 로 거절당해 조회수가 통째로 빈칸이 된다.
        # 엑셀에서 넘어온 BOM(U+FEFF)도 같은 이유로 제거한다. 표시용 키워드(kw)는 그대로 둔다.
        hint = kw.replace("\ufeff", "").replace(" ", "").replace("\u3000", "").strip()
        if not hint:
            result[kw] = {"pc": None, "mobile": None, "related": None}
            continue
        ts = str(int(time.time() * 1000))
        sig = build_searchad_signature(ts, "GET", endpoint, secret_key)
        headers = {"X-Timestamp": ts, "X-API-KEY": api_key, "X-Customer": customer_id, "X-Signature": sig}
        params = {"hintKeywords": hint, "showDetail": 1}
        try:
            r = requests.get("https://api.searchad.naver.com" + endpoint, headers=headers, params=params, timeout=30)
            if r.status_code >= 400:
                print(f"[조회수] 실패 {r.status_code}: {kw} (보낸 값: {hint})")
                result[kw] = {"pc": None, "mobile": None, "related": None}
                continue
            items = r.json().get("keywordList") or []
            picked = next((x for x in items if (x.get("relKeyword") or "") == hint), None) or (items[0] if items else None)
            if not picked:
                result[kw] = {"pc": 0, "mobile": 0, "related": kw}
                continue
            result[kw] = {"pc": parse_count(picked.get("monthlyPcQcCnt")), "mobile": parse_count(picked.get("monthlyMobileQcCnt")), "related": picked.get("relKeyword") or kw}
        except Exception as e:
            print(f"[조회수] 예외: {kw} ({type(e).__name__}: {e})")
            result[kw] = {"pc": None, "mobile": None, "related": None}
    return result


def api_search(client_id: str, client_secret: str, endpoint: str, query: str) -> list[dict[str, Any]] | None:
    """네이버 검색 API 호출. 일시적 실패(rate limit/5xx/네트워크)는 3회까지 재시도.
    - 200 OK: items 반환 (없으면 빈 리스트)
    - 4xx (429 제외): 영구 실패로 간주, 즉시 None
    - 429/5xx/네트워크 예외: 백오프 후 재시도
    """
    url = f"https://openapi.naver.com/v1/search/{endpoint}.json"
    headers = {"X-Naver-Client-Id": client_id, "X-Naver-Client-Secret": client_secret}
    params = (
        {"query": query, "display": 5, "sort": "random"}
        if endpoint == "local"
        else {"query": query, "display": 100, "sort": "sim"}
    )
    for attempt in range(3):
        try:
            r = requests.get(url, headers=headers, params=params, timeout=30)
            if r.status_code == 200:
                return r.json().get("items") or []
            if r.status_code != 429 and 400 <= r.status_code < 500:
                # 클라이언트 에러(잘못된 쿼리/인증 등)는 재시도해도 동일 — 즉시 실패
                return None
        except Exception:
            pass
        if attempt < 2:
            time.sleep(0.8 * (2 ** attempt))  # 0.8s, 1.6s
    return None


def item_text_for_tab(tab: str, item: dict[str, Any]) -> str:
    if tab == "blog":
        return " ".join([strip_html(item.get("bloggername", "")), strip_html(item.get("title", "")), strip_html(item.get("description", ""))])
    if tab == "cafe":
        return " ".join([strip_html(item.get("name", "")), strip_html(item.get("nickname", "")), strip_html(item.get("title", "")), strip_html(item.get("description", ""))])
    if tab == "map":
        return " ".join([strip_html(item.get("title", "")), strip_html(item.get("category", "")), strip_html(item.get("address", "")), strip_html(item.get("roadAddress", ""))])
    if tab == "news":
        return " ".join([strip_html(item.get("title", "")), strip_html(item.get("description", "")), strip_html(item.get("originallink", ""))])
    if tab == "video":
        return " ".join(
            [
                strip_html(item.get("title", "")),
                strip_html(item.get("description", "")),
                strip_html(item.get("author", "")),
            ]
        )
    if tab == "web":
        return " ".join([strip_html(item.get("title", "")), strip_html(item.get("description", "")), strip_html(item.get("link", ""))])
    return ""


def analyze_blog_search_items(
    items: list[dict[str, Any]],
    match_tokens: list[str],
    blog_period: tuple[int, int] | None,
    official_blog_ids: frozenset[str] = frozenset(),
) -> tuple[int, dict[str, Any]]:
    """블로그 API items: 1~10위는 점수용, 11~100위는 evidence(debug)용.

    점수 규칙: 1~10위 중 hospitalBlogBases에 등록된 공식 네이버 블로그 ID와 일치하는 첫 글의 rank로 채점.
    글 본문/제목의 키워드 관련성은 점수에 사용하지 않음 (matchTitleOrBodyOnly evidence로만 표시).
    """
    top: list[dict[str, Any]] = []
    matched = 0
    extra: dict[str, Any] = {
        "blogMatchRule": "official_blog_url_whitelist_only",
        "verifySources": {"openapi_top10": {"checked": True}},
    }
    if official_blog_ids:
        extra["officialNaverBlogIds"] = sorted(official_blog_ids)
    if blog_period:
        extra["blogScoringPeriod"] = {"year": blog_period[0], "month": blog_period[1]}
    for i, it in enumerate(items[:10], start=1):
        full_txt = item_text_for_tab("blog", it)
        auth_txt = blog_author_blog_text_for_match(it)
        official_ok = blog_item_matches_official_naver_blog(it, official_blog_ids)
        token_ok = tokens_match_in_normalized(auth_txt, match_tokens)
        body_ok = tokens_match_in_normalized(full_txt, match_tokens)
        if blog_period:
            in_m = blog_item_in_scoring_month(it, blog_period[0], blog_period[1])
        else:
            in_m = True
        row: dict[str, Any] = {
            "rank": i,
            "text": full_txt[:220],
            "authorBlogMatchText": auth_txt[:280],
            "postdate": it.get("postdate"),
            "blogInScoringMonth": in_m,
            "matchOfficialNaverBlog": official_ok,
            "matchAuthorOrBlogTokens": token_ok,
            "matchAuthorOrBlog": official_ok,
            "matchTitleOrBodyOnly": bool(body_ok and not official_ok),
            "bloggername": strip_html(it.get("bloggername", "")),
            "bloggerlink": strip_html(it.get("bloggerlink", "")),
        }
        top.append(row)
        if matched == 0 and in_m and official_ok:
            matched = i
            extra["matched_text"] = auth_txt[:280]
            extra["matched_postdate"] = it.get("postdate")
            extra["verifySources"]["openapi_top10"]["matched_rank"] = i
            extra["matchedVia"] = "official_hospital_blog_url"
    # 11~100위는 evidence(debug)용 — 점수 부여 X, 매칭 후보만 기록
    extended_matches: list[dict[str, Any]] = []
    for i, it in enumerate(items[10:100], start=11):
        auth_txt = blog_author_blog_text_for_match(it)
        official_ok = blog_item_matches_official_naver_blog(it, official_blog_ids)
        token_ok = tokens_match_in_normalized(auth_txt, match_tokens)
        if official_ok:
            extended_matches.append({
                "rank": i,
                "postdate": it.get("postdate"),
                "link": it.get("link"),
                "bloggername": strip_html(it.get("bloggername", "")),
                "matchOfficialNaverBlog": official_ok,
                "matchAuthorOrBlogTokens": token_ok,
            })
    if extended_matches:
        extra["verifySources"]["openapi_top11to100_debug"] = {
            "scoring": False,  # 점수에 사용 X, 디버깅 참고용
            "matches": extended_matches[:20],
            "matchCount": len(extended_matches),
        }
    out: dict[str, Any] = {"top": top, "matched_rank": matched, **extra}
    if matched == 0:
        body_only_in_month = False
        official_out_month = False
        for it in items[:10]:
            full_txt = item_text_for_tab("blog", it)
            official_ok = blog_item_matches_official_naver_blog(it, official_blog_ids)
            body_ok = tokens_match_in_normalized(full_txt, match_tokens)
            in_m = blog_item_in_scoring_month(it, blog_period[0], blog_period[1]) if blog_period else True
            if in_m and body_ok and not official_ok:
                body_only_in_month = True
            if blog_period and not in_m and official_ok:
                official_out_month = True
        if not official_blog_ids:
            out["blogNote"] = "hospitalBlogBases에 공식 네이버 블로그 ID가 등록되지 않아 블로그 0점 처리."
        elif body_only_in_month:
            out["blogNote"] = (
                "상위 노출 중 제목·본문에만 병원명이 있고, hospitalBlogBases에 등록된 공식 블로그 ID와 맞지 않아 블로그 0점 처리."
            )
        elif official_out_month:
            out["blogNote"] = (
                "공식 블로그 ID 기준은 맞으나 네이버 표기 작성일(postdate)이 "
                f"{blog_period[0]}년 {blog_period[1]}월이 아님."
            )
        elif blog_period:
            out["blogNote"] = (
                f"배점 월({blog_period[0]}년 {blog_period[1]}월) 내 작성이면서 공식 블로그 ID 기준에 맞는 글이 없음."
            )
    return (matched if matched else 0), out


def find_official_blog_post_for_keyword(
    keyword: str,
    hospital_names: list[str],
    official_blog_ids: frozenset[str],
    client_id: str,
    client_secret: str,
) -> dict[str, Any] | None:
    """공식 블로그가 특정 키워드 관련 글을 보유하는지 빠르게 확인.

    원리: 단순 `query=keyword`로는 openapi blog가 공식 블로그 글을 안 잡는 경우가 흔하지만
    (예: 글 제목·본문에 키워드가 정확히 안 들어간 경우),
    `query=keyword + 병원명`으로 결합하면 같은 글이 색인에서 잡힌다.

    반환: 공식 블로그 글 발견 시 {"link", "postdate", "title", "matched_via_query"}, 못 찾으면 None.
    Playwright verify trigger 및 evidence 용도 — 점수 부여 X.
    """
    if not official_blog_ids or not hospital_names:
        return None
    for name in hospital_names:
        name_clean = (name or "").strip()
        if not name_clean:
            continue
        q = f"{keyword} {name_clean}"
        items = api_search(client_id, client_secret, "blog", q) or []
        for it in items:
            link = (it.get("link") or "").strip()
            bid = naver_blog_id_from_url(link) or naver_blog_id_from_url(it.get("bloggerlink") or "")
            if bid and bid in official_blog_ids:
                return {
                    "link": link,
                    "postdate": it.get("postdate"),
                    "title": strip_html(it.get("title") or ""),
                    "matched_via_query": q,
                    "matched_blog_id": bid,
                }
    return None


def find_rank_by_api_tab(
    tab: str,
    query: str,
    match_tokens: list[str],
    client_id: str,
    client_secret: str,
    blog_period: tuple[int, int] | None = None,
    official_blog_ids: frozenset[str] = frozenset(),
    official_cafe_ids: frozenset[str] = frozenset(),
    device: str = "pc",
) -> tuple[int | None, dict[str, Any]]:
    endpoint = TAB_ENDPOINT.get(tab)
    if not endpoint:
        return None, {"reason": "unsupported endpoint"}
    items = api_search(client_id, client_secret, endpoint, query)
    if items is None:
        # map: 공식 지역 API 호출 자체 실패해도 통합검색 plat 카드(drt 메타)에서 fallback 매칭 시도
        if tab == "map":
            fb_rank, fb_ev = _try_map_drt_fallback(query, match_tokens, primary_basis="api_error", device=device)
            if fb_rank > 0:
                return fb_rank, fb_ev
            return None, {"reason": "api_error", **fb_ev}
        return None, {"reason": "api_error"}
    if tab == "blog":
        return analyze_blog_search_items(items, match_tokens, blog_period, official_blog_ids)
    if tab == "cafe":
        top: list[dict[str, Any]] = []
        matched = 0
        extra: dict[str, Any] = {
            "cafeMatchRule": "official_cafe_url_and_author_is_hospital_in_scoring_month",
        }
        if official_cafe_ids:
            extra["officialNaverCafeIds"] = sorted(official_cafe_ids)
        if blog_period:
            extra["cafeScoringPeriod"] = {"year": blog_period[0], "month": blog_period[1]}
        for i, it in enumerate(items[:10], start=1):
            txt = item_text_for_tab("cafe", it)
            official_ok = cafe_item_matches_official_naver_cafe(it, official_cafe_ids)
            nickname_txt = normalize_text(strip_html(str(it.get("nickname", ""))))
            author_ok = tokens_match_in_normalized(nickname_txt, match_tokens)
            in_m = cafe_item_in_scoring_month(it, blog_period[0], blog_period[1]) if blog_period else True
            row: dict[str, Any] = {
                "rank": i,
                "text": txt[:220],
                "date": it.get("date"),
                "matchOfficialNaverCafe": official_ok,
                "matchAuthorIsHospital": author_ok,
                "cafeInScoringMonth": in_m,
                "nickname": strip_html(str(it.get("nickname", ""))),
            }
            top.append(row)
            if matched == 0 and official_ok and author_ok and in_m:
                matched = i
                extra["matched_text"] = txt[:220]
                extra["matched_date"] = it.get("date")
                extra["matchedVia"] = "official_cafe_url_and_author_in_month"
        return (matched if matched else 0), {"top": top, "matched_rank": matched, **extra}
    has_date_filter = tab in ("news", "video") and blog_period is not None
    if has_date_filter:
        extra: dict[str, Any] = {"scoringPeriod": {"year": blog_period[0], "month": blog_period[1]}}
    else:
        extra = {}
    top: list[dict[str, Any]] = []
    matched = 0
    for i, it in enumerate(items[:10], start=1):
        txt = item_text_for_tab(tab, it)
        in_m = content_item_in_scoring_month(tab, it, blog_period[0], blog_period[1]) if has_date_filter else True
        date_val = it.get("pubDate") if tab == "news" else it.get("date")
        row: dict[str, Any] = {"rank": i, "text": txt[:220]}
        if has_date_filter:
            row["date"] = date_val
            row["inScoringMonth"] = in_m
        top.append(row)
        if matched == 0 and in_m and any(n in normalize_text(txt) for n in match_tokens):
            matched = i
            extra["matched_text"] = txt[:220]
            if has_date_filter:
                extra["matched_date"] = date_val
    # map: 공식 지역 API top10에 매칭 없을 때 통합검색 plat 카드(drt 메타) fallback
    if tab == "map" and matched == 0:
        fb_rank, fb_ev = _try_map_drt_fallback(query, match_tokens, primary_basis="api_top10_no_match", device=device)
        if fb_rank > 0:
            return fb_rank, {**fb_ev, "primaryApiTop": top}
        extra["drtFallback"] = fb_ev
    return (matched if matched else 0), {"top": top, "matched_rank": matched, **extra}


def fetch_search_page(query: str, where: str | None = None, device: str = "pc") -> str | None:
    host = "m.search.naver.com" if device == "mobile" else "search.naver.com"
    base = f"https://{host}/search.naver?query=" + quote_plus(query)
    urls = [base]
    if where:
        urls.insert(0, base + "&where=" + quote_plus(where))
    session = _get_web_session(device)
    for _ in range(3):
        for url in urls:
            try:
                r = session.get(url, timeout=30)
                if r.status_code >= 400:
                    continue
                time.sleep(random.uniform(1.5, 3.5))
                return decode_response_text(r)
            except Exception:
                continue
        time.sleep(random.uniform(1.5, 3.5))
    return None


_INTEGRATED_HTML_CACHE: dict[str, str] = {}
_INTEGRATED_HTML_CACHE_LOCK = threading.Lock()
_INTEGRATED_HTML_CACHE_MAX = 64


def fetch_integrated_search_page(query: str, device: str = "pc") -> str | None:
    """통합검색 페이지 fetch. 같은 query는 thread-safe 캐시(maxsize=64)로 재사용.

    map(drt fallback) / web / video fallback / powerlink·bizsite fallback이
    같은 query의 통합검색 HTML을 공유. 추가 fetch를 0회로 줄여 throttle 영향 최소화.
    fetch 실패(None)는 캐시하지 않아 다음 호출에서 재시도 가능.
    """
    cache_key = f"{device}:{query}"
    with _INTEGRATED_HTML_CACHE_LOCK:
        cached = _INTEGRATED_HTML_CACHE.get(cache_key)
    if cached is not None:
        return cached

    host = "m.search.naver.com" if device == "mobile" else "search.naver.com"
    urls = [
        f"https://{host}/search.naver?where=nexearch&sm=tab_jum&ssc=tab.nx.all&query=" + quote_plus(query),
        f"https://{host}/search.naver?query=" + quote_plus(query),
    ]
    session = _get_web_session(device)
    for _ in range(3):
        for url in urls:
            try:
                r = session.get(url, timeout=30)
                if r.status_code >= 400:
                    continue
                time.sleep(random.uniform(1.5, 3.5))
                html_text = decode_response_text(r)
                if html_text:
                    with _INTEGRATED_HTML_CACHE_LOCK:
                        if len(_INTEGRATED_HTML_CACHE) >= _INTEGRATED_HTML_CACHE_MAX:
                            # 단순 LRU: 가장 오래된 항목 1개 제거
                            _INTEGRATED_HTML_CACHE.pop(next(iter(_INTEGRATED_HTML_CACHE)))
                        _INTEGRATED_HTML_CACHE[cache_key] = html_text
                return html_text
            except Exception:
                continue
        time.sleep(random.uniform(1.5, 3.5))
    return None


_DRT_META_RE = re.compile(r'"id":"(\d+)","dbType":"drt","name":"([^"]+)"')


def _try_map_drt_fallback(
    query: str,
    match_tokens: list[str],
    *,
    primary_basis: str,
    device: str = "pc",
) -> tuple[int, dict[str, Any]]:
    """
    NAVER 공식 지역검색 API top10 매칭 실패 시 통합검색 페이지의 플레이스 카드(drt 메타) fallback.

    NAVER 공식 지역검색 API(local.json)는 통상 5건만 반환하며 거리·텍스트 매칭 기반.
    통합검색 페이지의 플레이스 카드는 위치/카테고리/리뷰 종합 추천 결과로 다른 hospital list가 노출될 수 있다.
    사용자가 브라우저에서 보는 화면 = 통합검색 카드라 0점 미스매치를 보강한다.

    drt 메타는 통상 8건. 그 안에 매칭되면 해당 순위(1~8) 부여, 없으면 0점.
    """
    ht = fetch_integrated_search_page(query, device)
    if not ht:
        return 0, {
            "matched_rank": 0,
            "basis": "drt_fallback_fetch_failed",
            "primaryBasis": primary_basis,
        }
    metas = _DRT_META_RE.findall(ht)
    top: list[dict[str, Any]] = []
    matched = 0
    for i, (pid, name) in enumerate(metas[:8], start=1):
        ntxt = normalize_text(name)
        is_match = any(t in ntxt for t in match_tokens)
        top.append({"rank": i, "name": name, "id": pid, "match": is_match})
        if matched == 0 and is_match:
            matched = i
    return matched, {
        "matched_rank": matched,
        "basis": "integrated_search_drt_fallback",
        "drtTop": top,
        "drtCount": len(metas),
        "primaryBasis": primary_basis,
        "device": device,
    }


def fetch_powerlink_more_page(query: str, device: str = "pc") -> str | None:
    """
    파워링크는 통합검색 메인 블록이 아닌 '더보기(광고 전체)' 기준으로 순위를 산정한다.
    """
    url = "https://ad.search.naver.com/search.naver?where=ad&query=" + quote_plus(query)
    session = _get_web_session(device)
    for _ in range(3):
        try:
            r = session.get(url, timeout=30)
            if r.status_code >= 400:
                time.sleep(random.uniform(1.5, 3.5))
                continue
            time.sleep(random.uniform(1.5, 3.5))
            enc = (r.apparent_encoding or r.encoding or "utf-8").strip()
            try:
                return r.content.decode(enc, errors="ignore")
            except Exception:
                return r.text
        except Exception:
            time.sleep(random.uniform(1.5, 3.5))
            continue
    return None


def extract_candidates_powerlink(ht: str) -> list[str]:
    soup = BeautifulSoup(ht, "html.parser")
    root = soup.select_one(
        "div[id^='pcPowerLink_'], div[id^='moPowerLink_'], "
        "div[id*='PowerLink_'], section[class*='powerlink']"
    )
    if not root:
        return []
    vals = []
    for n in root.select("li, .ad_dsc, .url_area, .desc"):
        t = strip_html(n.get_text(" ", strip=True))
        if t:
            vals.append(t)
    return vals[:20]


_POWERLINK_AD_SELECTOR = "li[data-index], li.lst, ul.lst_type > li, div.ad_list li"


def extract_candidates_powerlink_more(ht: str) -> list[str]:
    """
    ad.search.naver.com(where=ad) 페이지의 광고 목록을 DOM document order로 추출한다.

    여러 selector를 CSS ',' 결합으로 한 번에 매칭하면 BeautifulSoup이 document order로
    유니크 매칭을 반환하므로, 단일 selector가 일부만 잡는 경우에도 누락 없이 보강된다.
    `ol > li`는 광고 외 항목 섞일 위험으로 제외.
    """
    soup = BeautifulSoup(ht, "html.parser")
    vals: list[str] = []
    seen_ids: set[int] = set()
    for n in soup.select(_POWERLINK_AD_SELECTOR):
        if id(n) in seen_ids:
            continue
        seen_ids.add(id(n))
        t = strip_html(n.get_text(" ", strip=True))
        if t and len(t) > 3:
            vals.append(t)
    return vals[:30]


def extract_candidates_bizsite(ht: str) -> list[str]:
    """
    네이버 비즈사이트(유료 사이트 노출) 후보 추출.
    네이버가 템플릿을 여러 번 교체해 왔어서 4단계로 폴백한다:
      1) data-block-id 에 'bizsite' 또는 'site/' 가 포함된 블록
      2) 구형 section class (sp_nsite / sp_nbizsite / sp_nbiz)
      3) section class 에 'bizsite' / 'site' 가 포함된 신형 패턴
      4) 제목(h2/h3)이 '비즈사이트' 또는 '사이트'인 섹션
    """
    soup = BeautifulSoup(ht, "html.parser")
    vals: list[str] = []

    # 1) data-block-id 기반 (가장 최근 템플릿 전환 대응)
    for block in soup.select("[data-block-id]"):
        bid = (block.get("data-block-id") or "").lower()
        if not bid:
            continue
        if ("bizsite" in bid) or bid.startswith("site/") or bid.startswith("ugs_site") or bid.startswith("ups_site"):
            nodes = block.select("li.bx") or block.select("li") or [block]
            for n in nodes:
                t = strip_html(n.get_text(" ", strip=True))
                if t and len(t) > 3:
                    vals.append(t)
            if vals:
                return vals[:20]

    # 2) 구형 section class
    for sel in ["section.sp_nsite li.bx", "section.sp_nbizsite li.bx", "section.sp_nbiz li.bx"]:
        nodes = soup.select(sel)
        if nodes:
            for n in nodes:
                t = strip_html(n.get_text(" ", strip=True))
                if t:
                    vals.append(t)
            if vals:
                return vals[:20]

    # 3) 신형 section class 패턴 매칭
    for sel in [
        'section[class*="bizsite"] li.bx',
        'section[class*="nbizsite"] li.bx',
        'section[class*="sp_site"] li.bx',
        'div[class*="bizsite"] li.bx',
    ]:
        nodes = soup.select(sel)
        if nodes:
            for n in nodes:
                t = strip_html(n.get_text(" ", strip=True))
                if t:
                    vals.append(t)
            if vals:
                return vals[:20]

    # 4) 헤딩 텍스트 폴백
    for section in soup.select("section, div.sc_new, div.api_subject_bx"):
        heading = section.find(["h2", "h3", "h4"])
        if not heading:
            continue
        htext = strip_html(heading.get_text(" ", strip=True))
        if not htext:
            continue
        if "비즈사이트" in htext or htext.strip() in ("사이트",):
            nodes = section.select("li.bx") or section.select("li")
            for n in nodes:
                t = strip_html(n.get_text(" ", strip=True))
                if t:
                    vals.append(t)
            if vals:
                return vals[:20]

    return vals[:20]


def extract_video_items_with_dates(ht: str) -> list[tuple[str, str | None]]:
    """동영상 탭 후보 목록 + 업로드 날짜 추출. (text, raw_date_str | None) 리스트 반환.
    text 는 채널명(author) + 영상 제목(title) 결합본 — 두 곳 어디에 병원명이 있어도 매칭되도록.
    """
    marker = '"blockId":"video/prs_template_v2_video_tab_desk.ts"'
    idx = ht.find(marker)
    if idx >= 0:
        chunk = ht[idx : idx + 500000]
        author_pat = re.compile(r'"authorHtml":"((?:\\.|[^"\\])*)"')
        author_matches = list(author_pat.finditer(chunk))
        if author_matches:
            date_pat = re.compile(r'"date":"((?:\\.|[^"\\])*)"')
            title_pat = re.compile(r'"title":"((?:\\.|[^"\\])*)"')
            date_matches = list(date_pat.finditer(chunk))
            title_matches = list(title_pat.finditer(chunk))
            result: list[tuple[str, str | None]] = []
            for i, am in enumerate(author_matches):
                try:
                    author = strip_html(json.loads('"' + am.group(1) + '"'))
                except Exception:
                    author = strip_html(am.group(1))
                if not author:
                    continue
                # 현재 authorHtml과 다음 authorHtml 사이에서 date 필드 탐색
                lo = am.start()
                hi = author_matches[i + 1].start() if i + 1 < len(author_matches) else lo + 4000
                date_raw: str | None = None
                for dm in date_matches:
                    if lo <= dm.start() < hi:
                        date_raw = dm.group(1)
                        break
                # 앞쪽 2000자 범위도 보조 탐색
                if date_raw is None:
                    for dm in date_matches:
                        if max(0, lo - 2000) <= dm.start() < lo:
                            date_raw = dm.group(1)
                # 같은 영역에서 영상 제목 추출
                title = ""
                for tm in title_matches:
                    if lo <= tm.start() < hi:
                        try:
                            title = strip_html(json.loads('"' + tm.group(1) + '"'))
                        except Exception:
                            title = strip_html(tm.group(1))
                        if title:
                            break
                text = f"{author} {title}".strip() if title else author
                result.append((text, date_raw))
            if result:
                return result[:30]
    # DOM 파싱 fallback (날짜 없음)
    soup = BeautifulSoup(ht, "html.parser")
    vals: list[tuple[str, str | None]] = []
    for sel in ["#main_pack li.bx", "section.sp_nvideo li.bx", "ul.lst_video li"]:
        nodes = soup.select(sel)
        if nodes:
            for n in nodes:
                t = strip_html(n.get_text(" ", strip=True))
                if t:
                    vals.append((t, None))
            if vals:
                break
    return vals[:30]


def extract_candidates_video(ht: str) -> list[str]:
    # 1) 동영상 탭 렌더 데이터(fender)에서 작성자/제목을 우선 추출
    marker = '"blockId":"video/prs_template_v2_video_tab_desk.ts"'
    idx = ht.find(marker)
    if idx >= 0:
        chunk = ht[idx : idx + 500000]
        # 작성자 기준 매칭이 핵심이므로 authorHtml 우선 추출
        author_pat = re.compile(r'"authorHtml":"((?:\\.|[^"\\])*)"')
        authors: list[str] = []
        for m in author_pat.finditer(chunk):
            try:
                author = strip_html(json.loads('"' + m.group(1) + '"'))
            except Exception:
                author = strip_html(m.group(1))
            if author:
                authors.append(author)
        if authors:
            return authors[:30]

        pat = re.compile(
            r'"authorHtml":"((?:\\.|[^"\\])*)".{0,2600}?"title":"((?:\\.|[^"\\])*)"',
            re.S,
        )
        parsed: list[str] = []
        for m in pat.finditer(chunk):
            try:
                author = strip_html(json.loads('"' + m.group(1) + '"'))
            except Exception:
                author = strip_html(m.group(1))
            try:
                title = strip_html(json.loads('"' + m.group(2) + '"'))
            except Exception:
                title = strip_html(m.group(2))
            text = f"{author} {title}".strip()
            if text:
                parsed.append(text)
        if parsed:
            return parsed[:30]

    # 2) 일반 DOM 파싱 fallback
    soup = BeautifulSoup(ht, "html.parser")
    vals = []
    for sel in ["#main_pack li.bx", "section.sp_nvideo li.bx", "ul.lst_video li"]:
        nodes = soup.select(sel)
        if nodes:
            for n in nodes:
                t = strip_html(n.get_text(" ", strip=True))
                if t:
                    vals.append(t)
            if vals:
                break
    return vals[:30]


def extract_candidates_web_from_integrated(ht: str) -> list[str]:
    """
    웹 점수 기준:
    - 네이버 통합검색 첫 화면에서 '블로그 성격 결과'를 후보로 추출
    - 해당 후보의 순위를 web 열 점수 산정에 사용
    """
    # 1) 통합검색 페이지의 블로그 섹션 DOM 우선
    soup = BeautifulSoup(ht, "html.parser")
    vals: list[str] = []
    # 0) data-block-id별 블로그·UGC 묶음 (네이버가 템플릿 id를 바꾸면 여기에 추가)
    for bid in (
        "web/prs_template_v2_web_basic_desk.ts",
        "review/prs_template_v2_review_blog_rra_desk.ts",
        "review/prs_template_v2_review_ugc_single_intention_desk.ts",
        "review/prs_template_v2_review_ugc_single_intention_mob.ts",
    ):
        block = soup.select_one(f"div[data-block-id='{bid}']")
        if not block:
            continue
        for item_sel in ("li.bx", "ul.lst_view li", "li"):
            nodes = block.select(item_sel)
            texts: list[str] = []
            for n in nodes:
                t = strip_html(n.get_text(" ", strip=True))
                if t and len(t) > 12:
                    texts.append(t)
            if len(texts) >= 2:
                return texts[:30]
        t = strip_html(block.get_text(" ", strip=True))
        if t:
            return [t]

    for sel in [
        "section._sp_nblog li.bx",
        "section.sp_nblog li.bx",
        "#main_pack section._sp_nblog li.bx",
    ]:
        nodes = soup.select(sel)
        if nodes:
            for n in nodes:
                t = strip_html(n.get_text(" ", strip=True))
                if t:
                    vals.append(t)
            if vals:
                return vals[:30]

    # 2) 렌더 데이터 fallback — 신형 web_basic이 페이지에 있으면 우선
    idx = -1
    for marker in (
        "web/prs_template_v2_web_basic_desk.ts",
        "review/prs_template_v2_review_blog_rra_desk.ts",
        "review/prs_template_v2_review_blog_tab_desk.ts",
        "review/prs_template_v2_review_ugc_single_intention_desk.ts",
        "review/prs_template_v2_review_ugc_single_intention_mob.ts",
    ):
        i = ht.find(marker)
        if i >= 0:
            idx = i
            break
    if idx >= 0:
        chunk = ht[idx : idx + 450000]
        author_pat = re.compile(r'"authorHtml":"((?:\\.|[^"\\])*)"')
        authors: list[str] = []
        for m in author_pat.finditer(chunk):
            try:
                author = strip_html(json.loads('"' + m.group(1) + '"'))
            except Exception:
                author = strip_html(m.group(1))
            if author:
                authors.append(author)
        if authors:
            return authors[:30]

        pat = re.compile(
            r'"authorHtml":"((?:\\.|[^"\\])*)".{0,2600}?"title":"((?:\\.|[^"\\])*)"',
            re.S,
        )
        parsed: list[str] = []
        for m in pat.finditer(chunk):
            try:
                author = strip_html(json.loads('"' + m.group(1) + '"'))
            except Exception:
                author = strip_html(m.group(1))
            try:
                title = strip_html(json.loads('"' + m.group(2) + '"'))
            except Exception:
                title = strip_html(m.group(2))
            text = f"{author} {title}".strip()
            if text:
                parsed.append(text)
        if parsed:
            return parsed[:30]
    return []


def find_rank_in_candidates(cands: list[str], match_tokens: list[str]) -> int:
    for i, t in enumerate(cands[:10], start=1):
        if any(n in normalize_text(t) for n in match_tokens):
            return i
    return 0


def find_extended_rank_in_candidates(cands: list[str], match_tokens: list[str]) -> int:
    """11~20위 매칭. evidence/debug 전용 — 점수 부여 금지."""
    for i, t in enumerate(cands[10:20], start=11):
        if any(n in normalize_text(t) for n in match_tokens):
            return i
    return 0


def _find_web_rank_by_url(
    ht: str,
    match_tokens: list[str],
    official_blog_ids: frozenset[str] = frozenset(),
) -> tuple[int, dict[str, Any]]:
    """
    통합검색 HTML의 결과 링크(href)에서 병원 도메인 또는 공식 블로그 ID를
    직접 찾아 순위를 반환한다. 텍스트 내 언급(mention)은 무시한다.
    """
    soup = BeautifulSoup(ht, "html.parser")
    main = soup.select_one("#main_pack") or soup

    seen: set[str] = set()
    result_urls: list[str] = []
    for a in main.find_all("a", href=True):
        href = (a.get("href") or "").strip()
        if not href or href.startswith("#") or "javascript" in href.lower():
            continue
        # 절대 URL(http/https)만 외부 결과로 인정 — 상대경로 검색옵션 링크(?ssc=...) 제외
        if not href.lower().startswith(("http://", "https://")):
            continue
        # 광고 및 네이버 내부 검색·네비·도움말 링크 제외
        if any(s in href for s in (
            "ad.search.naver.com",
            "link.naver.com",
            "search.naver.com/search",
            "ader.naver.com",
            "help.naver.com",
            "saedu.naver.com",
        )):
            continue
        if href not in seen:
            seen.add(href)
            result_urls.append(href)

    top = [u[:150] for u in result_urls[:15]]

    # 도메인 토큰: "blog.naver.com" 같은 네이버 공용 도메인과 5자 미만 토큰은 제외
    domain_tokens = [
        t for t in match_tokens
        if len(t) >= 5 and "naver.com" not in t and "." in t
    ]

    for rank, url in enumerate(result_urls, start=1):
        url_l = url.lower()
        for t in domain_tokens:
            if t in url_l:
                return rank, {"matched_url": url[:200], "matched_rank": rank, "basis": "web_url_domain", "top": top}
        for bid in official_blog_ids:
            if text_has_naver_blog_id(url_l, bid):
                return rank, {"matched_url": url[:200], "matched_rank": rank, "basis": "web_url_blog", "top": top}

    return 0, {"matched_rank": 0, "basis": "no_url_match", "top": top}


def _find_web_rank_from_render_json(
    ht: str,
    match_tokens: list[str],
    official_blog_ids: frozenset[str] = frozenset(),
) -> tuple[int, dict[str, Any]]:
    """
    Naver HTML에 삽입된 VIEW 섹션 렌더 JSON에서 링크를 추출해 URL 기반 매칭.
    _find_web_rank_by_url()의 정적 href 매칭이 0을 반환할 때 fallback으로 사용.
    JS 렌더링으로 <a href>에 나타나지 않는 블로그 링크를 잡아낸다.
    """
    VIEW_MARKERS = [
        "web/prs_template_v2_web_basic_desk.ts",
        "review/prs_template_v2_review_blog_rra_desk.ts",
        "review/prs_template_v2_review_blog_tab_desk.ts",
        "review/prs_template_v2_review_ugc_single_intention_desk.ts",
        "review/prs_template_v2_review_ugc_single_intention_mob.ts",
    ]
    # 여러 마커가 동시에 있을 수 있으므로 가장 빠른 위치부터 한 번에 큰 청크를 잡는다.
    earliest = -1
    for marker in VIEW_MARKERS:
        idx = ht.find(marker)
        if idx >= 0 and (earliest < 0 or idx < earliest):
            earliest = idx
    if earliest < 0:
        return 0, {"matched_rank": 0, "basis": "no_view_section_in_json"}
    chunk = ht[earliest: earliest + 600000]

    # 신형 web_basic 블록은 titleHref/contentHref/imageHref/href 키 사용,
    # 구형 review/blog 블록은 link/blogLink/postUrl/mobileLink/url 키 사용.
    link_pat = re.compile(
        r'"(?:link|blogLink|postUrl|mobileLink|url|titleHref|contentHref|imageHref|href)"\s*:\s*"(https?://(?:[^"\\]|\\.)+)"'
    )
    domain_tokens = [
        t for t in match_tokens
        if len(t) >= 5 and "naver.com" not in t and "." in t
    ]

    seen: set[str] = set()
    extracted: list[str] = []
    for m in link_pat.finditer(chunk):
        url = m.group(1).replace("\\/", "/").replace("\\u002F", "/")
        if url not in seen:
            seen.add(url)
            extracted.append(url)

    top = [u[:150] for u in extracted[:15]]

    for rank, url in enumerate(extracted, start=1):
        url_l = url.lower()
        for t in domain_tokens:
            if t in url_l:
                return rank, {
                    "matched_url": url[:200], "matched_rank": rank,
                    "basis": "web_render_json_domain", "top": top,
                }
        for bid in official_blog_ids:
            if text_has_naver_blog_id(url_l, bid):
                return rank, {
                    "matched_url": url[:200], "matched_rank": rank,
                    "basis": "web_render_json_blog", "top": top,
                }

    return 0, {"matched_rank": 0, "basis": "no_render_json_match", "top": top}


def find_rank_by_web_tab(
    tab: str,
    query: str,
    match_tokens: list[str],
    blog_period: tuple[int, int] | None = None,
    official_blog_ids: frozenset[str] = frozenset(),
    device: str = "pc",
) -> tuple[int | None, dict[str, Any]]:
    if tab == "powerlink":
        fetched_at = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        ht_more = fetch_powerlink_more_page(query, device)
        if ht_more:
            cands = extract_candidates_powerlink_more(ht_more)
            rank = find_rank_in_candidates(cands, match_tokens)
            top = [{"rank": i + 1, "text": t[:220]} for i, t in enumerate(cands[:10])]
            ev: dict[str, Any] = {
                "top": top,
                "matched_rank": rank,
                "basis": "powerlink_more",
                "extractedAdCount": len(cands),
                "extractionSelector": "union_dedupe",
                "fetchedAt": fetched_at,
                "surface": "dedicated_tab",
                "dedicatedTabRank": rank,
                "device": device,
            }
            ext_rank = find_extended_rank_in_candidates(cands, match_tokens)
            if ext_rank > 0:
                ev["extendedRank11to20"] = {
                    "rank": ext_rank,
                    "scoring": False,
                    "note": "11~20위 evidence-only. 점수 부여 대상 아님.",
                }
            return rank, ev
        # 더보기 페이지 실패 시 통합검색 페이지로 폴백 (캐시 활용)
        ht = fetch_integrated_search_page(query, device)
        if not ht:
            return None, {"reason": "http_error", "fetchedAt": fetched_at}
        cands = extract_candidates_powerlink(ht)
        rank = find_rank_in_candidates(cands, match_tokens)
        top = [{"rank": i + 1, "text": t[:220]} for i, t in enumerate(cands[:10])]
        return rank, {
            "top": top,
            "matched_rank": rank,
            "basis": "powerlink_main_fallback",
            "extractedAdCount": len(cands),
            "extractionSelector": f"integrated_{device}_powerlink",
            "surface": "integrated_search",
            "integratedRank": rank,
            "fetchedAt": fetched_at,
        }
    if tab == "web":
        # URL 기반 매칭: 결과 링크(href)에 병원 도메인/공식 블로그 ID가 직접 포함될 때만 점수 부여.
        # 텍스트에 병원명이 '언급'되는 경우는 제외 — 타 병원 비교 포스트 오매칭 방지.
        for _ in range(3):
            ht = fetch_integrated_search_page(query, device)
            if not ht:
                time.sleep(0.4)
                continue
            rank, ev = _find_web_rank_by_url(ht, match_tokens, official_blog_ids)
            if rank > 0:
                return rank, {**ev, "surface": "integrated_search", "integratedRank": rank, "device": device}
            rank, ev = _find_web_rank_from_render_json(ht, match_tokens, official_blog_ids)
            return rank, {**ev, "surface": "integrated_search", "integratedRank": rank, "device": device}
        return None, {"matched_rank": 0, "reason": "fetch_failed"}

    if tab == "bizsite":
        # 통합검색 페이지 사용 — 캐시로 web 채점과 공유
        ht = fetch_integrated_search_page(query, device)
        if not ht:
            return None, {"reason": "http_error"}
        cands = extract_candidates_bizsite(ht)
        rank = find_rank_in_candidates(cands, match_tokens)
        top = [{"rank": i + 1, "text": t[:220]} for i, t in enumerate(cands[:10])]
        return rank, {"top": top, "matched_rank": rank, "surface": "integrated_search", "integratedRank": rank, "device": device}

    if tab == "video":
        # video 별도 검색 페이지 — fetch 실패해도 0 반환 안 하고 통합검색 fallback로 이동
        ht_v = fetch_search_page(query, where="video", device=device)
        has_date_filter = blog_period is not None
        score_year, score_month = blog_period if blog_period else (0, 0)
        top_ev: list[dict[str, Any]] = []
        matched = 0
        if ht_v:
            items = extract_video_items_with_dates(ht_v)
            for i, (txt, date_raw) in enumerate(items[:10], start=1):
                row: dict[str, Any] = {"rank": i, "text": txt[:220]}
                if has_date_filter:
                    pd = parse_cafe_date(date_raw)
                    in_m = pd is not None and pd.year == score_year and pd.month == score_month
                    row["date"] = date_raw
                    row["inScoringMonth"] = in_m
                else:
                    in_m = True
                top_ev.append(row)
                if matched == 0 and in_m and any(n in normalize_text(txt) for n in match_tokens):
                    matched = i
        # video 페이지에서 매칭 실패 또는 fetch 실패 시 통합검색 페이지의 동영상 캐러셀 fallback
        if matched == 0:
            ht_int = fetch_integrated_search_page(query, device)
            if ht_int:
                items_int = extract_video_items_with_dates(ht_int)
                fb_top = [{"rank": i + 1, "text": t[:220]} for i, (t, _d) in enumerate(items_int[:10])]
                for i, (txt, _date) in enumerate(items_int[:10], start=1):
                    if any(n in normalize_text(txt) for n in match_tokens):
                        return i, {
                            "top": top_ev,
                            "matched_rank": i,
                            "basis": "integrated_search_video_fallback",
                            "primaryBasis": "video_page_fetch_failed" if not ht_v else "video_page_no_match",
                            "fallbackTop": fb_top,
                            "matched_text": txt[:220],
                            "surface": "integrated_search",
                            "dedicatedTabRank": matched,
                            "integratedRank": i,
                        }
        ev: dict[str, Any] = {
            "top": top_ev,
            "matched_rank": matched,
            "surface": "dedicated_tab",
            "dedicatedTabRank": matched,
        }
        if has_date_filter:
            ev["scoringPeriod"] = {"year": score_year, "month": score_month}
        if not ht_v and matched == 0:
            ev["videoPageFetchFailed"] = True
        return matched, ev

    return None, {"reason": "unsupported_web_tab"}


def build_row(row_no: int, region: str, keyword: str, points: dict[str, int | None], pc: int | None, mobile: int | None, related: str | None) -> list[Any]:
    row = [None] * 16
    row[0], row[1], row[2] = row_no, (region if region else None), keyword
    row[3], row[4], row[5], row[6] = None, pc, mobile, related
    for tab, col in COL_BY_TAB.items():
        row[col] = points.get(tab)
    total, has_any = 0, False
    for c in range(7, 15):
        if c == POWERLINK_COL:
            continue
        v = row[c]
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            total += int(v)
            has_any = True
    row[TOTAL_COL] = total if has_any else None
    return row


def load_config() -> dict[str, Any]:
    name = (os.getenv("SCORING_CONFIG") or "april_keywords.json").strip()
    path = ROOT / "config" / name
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def post_progress_webhook(**fields: Any) -> None:
    """채점 진행상황을 Render webhook으로 전송. 설정 없으면 조용히 무시."""
    url = (os.getenv("SCORING_WEBHOOK_URL") or "").strip()
    task_id = (os.getenv("SCORING_TASK_ID") or "").strip()
    if not url or not task_id:
        return
    body = {"taskId": task_id, **{k: v for k, v in fields.items() if v is not None}}
    secret = (os.getenv("SCORING_WEBHOOK_SECRET") or "").strip()
    headers = {"Content-Type": "application/json"}
    if secret:
        headers["X-Webhook-Secret"] = secret
    try:
        requests.post(url, json=body, headers=headers, timeout=10)
    except Exception as e:
        print(f"[webhook] 전송 실패 (무시): {e}")


def fetch_keyword_ranks(
    cfg: dict[str, Any],
    cid: str,
    csec: str,
    *,
    report_progress: bool = False,
    device: str = "pc",
    shared_api: tuple[dict[str, dict[str, int | None]], dict[str, Any]] | None = None,
    browser_verify: bool = True,
) -> tuple[dict[str, dict[str, int | None]], dict[str, Any]]:
    manual_by_tab = cfg.get("manualRanksByTab") or {}
    legacy_blog_manual = cfg.get("manualRanks") or {}
    names = cfg.get("hospitalNames") or []
    domains = cfg.get("hospitalDomains") or []
    match_tokens = build_match_tokens(names, domains)
    out: dict[str, dict[str, int | None]] = {}
    ev_all: dict[str, Any] = {}

    blog_period = blog_evidence_period(cfg)
    official_blog_ids = official_naver_blog_ids_from_config(cfg)
    official_cafe_ids = frozenset()  # 카페 채점은 전역 비활성화 (2026-04 이후)
    keywords_list = list(cfg.get("keywords") or [])
    total_keywords = len(keywords_list)
    progress_step = max(1, min(20, total_keywords // 10 if total_keywords >= 20 else 1))
    if report_progress:
        post_progress_webhook(
            status="running",
            message=f"{device.upper()} 채점 진행중 0/{total_keywords}",
            totalKeywords=total_keywords,
            processedKeywords=0,
            stage="scoring",
        )
    def _score_one_keyword(kw: str) -> tuple[str, dict, dict]:
        kw_out: dict[str, Any] = {}
        kw_ev: dict[str, Any] = {}
        for tab in COL_BY_TAB.keys():
            if tab == "cafe":
                kw_out[tab] = 0
                kw_ev[tab] = {
                    "source": "disabled",
                    "device": device,
                    "reason": "cafe_removed",
                    "matched_rank": 0,
                    "top": [],
                    "note": "카페 채점 전역 비활성화",
                }
                continue
            if shared_api and tab in ("blog", "news"):
                shared_ranks, shared_evidence = shared_api
                kw_out[tab] = (shared_ranks.get(kw) or {}).get(tab)
                shared_ev = dict((shared_evidence.get(kw) or {}).get(tab) or {})
                shared_ev["deviceScope"] = "shared_api"
                shared_ev["device"] = device
                shared_ev["surface"] = "dedicated_tab"
                shared_ev["dedicatedTabRank"] = shared_ev.get(
                    "matched_rank", kw_out[tab]
                )
                kw_ev[tab] = shared_ev
                continue
            if tab in manual_by_tab.get(kw, {}):
                r = int(manual_by_tab[kw][tab]); kw_out[tab] = r; kw_ev[tab] = {"source": "manual", "rank": r, "device": device}; continue
            if tab == "blog" and kw in legacy_blog_manual:
                r = int(legacy_blog_manual[kw]); kw_out[tab] = r; kw_ev[tab] = {"source": "manual_legacy_blog", "rank": r, "device": device}; continue
            if tab in ("powerlink", "bizsite", "video", "web"):
                r, ev = find_rank_by_web_tab(
                    tab, kw, match_tokens, blog_period, official_blog_ids, device=device
                )
                kw_out[tab] = r  # None = 측정 실패 → 표에 '—' (진짜 0점과 구분)
                kw_ev[tab] = {"source": "web", "device": device, **ev}
            elif tab == "map":
                # 지도는 지역 API 결과(최대 5건)보다 실제 통합검색 플레이스 화면을 우선한다.
                r, ev = _try_map_drt_fallback(
                    kw, match_tokens, primary_basis="integrated_first", device=device
                )
                if ev.get("basis") == "drt_fallback_fetch_failed":
                    r, ev = find_rank_by_api_tab(
                        tab, kw, match_tokens, cid, csec, device=device
                    )
                    kw_ev[tab] = {
                        "source": "api",
                        "device": device,
                        "surface": "dedicated_tab",
                        "dedicatedTabRank": r,
                        **ev,
                    }
                else:
                    kw_ev[tab] = {
                        "source": "web",
                        "device": device,
                        "surface": "integrated_search",
                        "integratedRank": r,
                        **ev,
                    }
                kw_out[tab] = r
            else:
                r, ev = find_rank_by_api_tab(
                    tab,
                    kw,
                    match_tokens,
                    cid,
                    csec,
                    blog_period if tab in ("blog", "cafe", "news", "video") else None,
                    official_blog_ids if tab == "blog" else frozenset(),
                    official_cafe_ids if tab == "cafe" else frozenset(),
                    device,
                )
                kw_out[tab] = r  # None = 측정 실패 → 표에 '—' (진짜 0점과 구분)
                kw_ev[tab] = {
                    "source": "api",
                    "device": device,
                    "surface": "dedicated_tab",
                    "dedicatedTabRank": r,
                    **ev,
                }
        return kw, kw_out, kw_ev

    max_workers = max(1, int(os.getenv("SCORING_PARALLEL_WORKERS", "3")))
    _lock = threading.Lock()
    _done = [0]

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(_score_one_keyword, kw): kw for kw in keywords_list}
        for future in as_completed(futures):
            kw = futures[future]
            try:
                kw, kw_out, kw_ev = future.result()
            except Exception as e:
                # 키워드 1개의 예외가 전체 채점을 중단시키지 않도록 격리한다.
                # 측정 실패는 0점이 아니라 "미측정"으로 남겨야 정확도 지표가 오염되지 않는다.
                # (kw_out 을 비우면 build_month_payload 에서 해당 키워드 점수가 None → 표에 '—' 로 표시)
                kw_out = {}
                kw_ev = {
                    tab: {"source": "error", "reason": "scoring_exception", "error": repr(e)}
                    for tab in COL_BY_TAB.keys()
                }
                print(f"  [경고] 키워드 채점 실패 — 건너뜀: {kw} ({e!r})", flush=True)
            out[kw] = kw_out
            ev_all[kw] = kw_ev
            if report_progress:
                with _lock:
                    _done[0] += 1
                    idx = _done[0]
                if idx == 1 or idx == total_keywords or idx % progress_step == 0:
                    pct = (idx * 100) // total_keywords if total_keywords else 0
                    print(f"  진행 {idx}/{total_keywords} ({pct}%) - {kw}", flush=True)
                    post_progress_webhook(
                        status="running",
                        message=f"{device.upper()} 채점 진행중 {idx}/{total_keywords}",
                        totalKeywords=total_keywords,
                        processedKeywords=idx,
                        stage="scoring",
                    )

    # Playwright 자동 교차검증 — 0점 의심 사례와 양수 표본을 모두 검사한다.
    # HTTP/API와 브라우저 결과가 다르면 브라우저를 한 번 더 실행해 다수결로 확정한다.
    verify_enabled = (os.getenv("PLAYWRIGHT_VERIFY_ENABLED") or "1").strip().lower() in {"1", "true", "yes"}
    if verify_enabled and browser_verify:
        integrated_channels = ("powerlink", "bizsite", "map", "video", "web")
        verify_keywords: list[str] = []
        blog_tab_keywords: set[str] = set()
        blog_official_post_found: dict[str, dict[str, Any]] = {}
        for kw, ch_ranks in out.items():
            blog_zero = (ch_ranks.get("blog") or 0) == 0
            if blog_zero and official_blog_ids and names:
                found = find_official_blog_post_for_keyword(
                    kw, names, official_blog_ids, cid, csec
                )
                if found:
                    blog_official_post_found[kw] = found
                    if kw not in verify_keywords:
                        verify_keywords.append(kw)
                    blog_tab_keywords.add(kw)

            ev_kw = ev_all.get(kw) or {}
            if any((ch_ranks.get(c) or 0) == 0 for c in integrated_channels):
                cache_key = f"{device}:{kw}"
                with _INTEGRATED_HTML_CACHE_LOCK:
                    ht_cached = _INTEGRATED_HTML_CACHE.get(cache_key)
                ht = ht_cached if ht_cached is not None else fetch_integrated_search_page(kw, device)
                if ht is not None and any(t in ht.lower() for t in match_tokens):
                    if kw not in verify_keywords:
                        verify_keywords.append(kw)
            if isinstance(ev_kw.get("web"), dict) and ev_kw["web"].get("reason") == "fetch_failed":
                if kw not in verify_keywords:
                    verify_keywords.append(kw)
            if isinstance(ev_kw.get("blog"), dict) and ev_kw["blog"].get("reason") == "api_error":
                if kw not in verify_keywords:
                    verify_keywords.append(kw)
                blog_tab_keywords.add(kw)

        # 양수 점수도 고정 표본으로 감사해 잘못 받은 점수를 탐지한다.
        audit_size = max(0, int(os.getenv("PLAYWRIGHT_AUDIT_SIZE", "12")))
        positive_pool = [
            kw for kw, ranks in out.items()
            if any((ranks.get(ch) or 0) > 0 for ch in integrated_channels)
        ]
        zero_pool = [
            kw for kw, ranks in out.items()
            if any((ranks.get(ch) or 0) == 0 for ch in integrated_channels)
        ]
        positive_pool.sort(
            key=lambda kw: hashlib.sha256(f"{device}:{kw}".encode("utf-8")).hexdigest()
        )
        zero_pool.sort(
            key=lambda kw: hashlib.sha256(f"zero:{device}:{kw}".encode("utf-8")).hexdigest()
        )
        for kw in positive_pool[:audit_size] + zero_pool[:audit_size]:
            if kw not in verify_keywords:
                verify_keywords.append(kw)

        if verify_keywords:
            print(
                f"\n[playwright_verify] {device.upper()} 자동 교차검증 {len(verify_keywords)}개 시작",
                flush=True,
            )
            if report_progress:
                post_progress_webhook(
                    status="running",
                    message=f"{device.upper()} 브라우저 교차검증 {len(verify_keywords)}개",
                    stage="playwright_verify",
                )
            try:
                from playwright_verify import verify_keywords as verify_with_browser
                first_results = verify_with_browser(
                    verify_keywords, match_tokens, official_blog_ids,
                    blog_tab_keywords=frozenset(blog_tab_keywords),
                    device=device,
                )

                def browser_rank(result: dict[str, Any], channel: str) -> int | None:
                    if not isinstance(result.get("__meta__"), dict) or not result["__meta__"].get("loaded"):
                        return None
                    info = result.get(channel)
                    return int(info.get("rank") or 0) if isinstance(info, dict) else 0

                disagreement_channels: dict[str, set[str]] = {}
                for kw in verify_keywords:
                    result = first_results.get(kw) or {}
                    for ch in integrated_channels:
                        br = browser_rank(result, ch)
                        if br is None:
                            continue
                        if table_cell_for_tab(ch, (out.get(kw) or {}).get(ch)) != table_cell_for_tab(ch, br):
                            disagreement_channels.setdefault(kw, set()).add(ch)
                disagreements = list(disagreement_channels)
                second_results = verify_with_browser(
                    list(dict.fromkeys(disagreements)),
                    match_tokens,
                    official_blog_ids,
                    device=device,
                ) if disagreements else {}

                corrected = 0
                uncertain = 0
                for kw in verify_keywords:
                    first = first_results.get(kw) or {}
                    second = second_results.get(kw) or {}
                    kw_out = out.get(kw) or {}
                    ev_kw = ev_all.setdefault(kw, {})
                    for ch in integrated_channels:
                        initial = kw_out.get(ch)
                        first_rank = browser_rank(first, ch)
                        if first_rank is None:
                            continue
                        second_rank = (
                            browser_rank(second, ch)
                            if ch in disagreement_channels.get(kw, set())
                            else first_rank
                        )
                        agreed = (
                            second_rank is not None
                            and table_cell_for_tab(ch, first_rank) == table_cell_for_tab(ch, second_rank)
                        )
                        final_rank = first_rank if agreed else None
                        ev_ch = ev_kw.get(ch) if isinstance(ev_kw.get(ch), dict) else {}
                        ev_ch["automationAudit"] = {
                            "device": device,
                            "primaryRank": initial,
                            "browserRank1": first_rank,
                            "browserRank2": second_rank,
                            "agreed": agreed,
                            "primaryAgreed": (
                                final_rank is not None
                                and table_cell_for_tab(ch, initial)
                                == table_cell_for_tab(ch, final_rank)
                            ),
                        }
                        if final_rank is None:
                            kw_out[ch] = None
                            ev_ch["preAuditMatchedRank"] = ev_ch.get("matched_rank")
                            ev_ch["matched_rank"] = None
                            ev_ch["reason"] = "automated_collectors_disagreed"
                            uncertain += 1
                        elif table_cell_for_tab(ch, initial) != table_cell_for_tab(ch, final_rank):
                            kw_out[ch] = final_rank
                            ev_ch["preAuditMatchedRank"] = ev_ch.get("matched_rank")
                            ev_ch["matched_rank"] = final_rank
                            ev_ch["source"] = "automated_consensus"
                            ev_ch["surface"] = "integrated_search"
                            ev_ch["integratedRank"] = final_rank
                            corrected += 1
                        ev_kw[ch] = ev_ch

                    # 블로그 통합검색 노출과 전용 탭 순위는 서로 덮어쓰지 않는다.
                    blog_info = first.get("blog")
                    if isinstance(blog_info, dict):
                        ev_blog = ev_kw.get("blog") if isinstance(ev_kw.get("blog"), dict) else {}
                        rank = int(blog_info.get("rank") or 0)
                        if blog_info.get("source") == "blog_tab_dom":
                            ev_blog["dedicatedTabVerifyRank"] = rank
                            ev_blog["integratedExposureRank"] = int(
                                blog_info.get("integratedExposureRank") or 0
                            )
                            if rank > 0:
                                kw_out["blog"] = rank
                                ev_blog["apiMatchedRank"] = ev_blog.get("matched_rank")
                                ev_blog["matched_rank"] = rank
                                ev_blog["source"] = "automated_consensus"
                                ev_blog["surface"] = "dedicated_tab"
                        else:
                            ev_blog["integratedExposureRank"] = rank
                        ev_kw["blog"] = ev_blog

                for kw, found in blog_official_post_found.items():
                    blog_rank = (out.get(kw) or {}).get("blog") or 0
                    if blog_rank == 0:
                        ev_kw = ev_all.setdefault(kw, {})
                        ev_blog = ev_kw.get("blog") or {}
                        if not isinstance(ev_blog, dict):
                            ev_blog = {"original": ev_blog}
                        ev_blog["automatic_retry_needed"] = True
                        ev_blog["officialBlogPostFound"] = found
                        ev_blog["automatic_retry_note"] = (
                            "공식 블로그가 이 키워드 관련 글을 보유하고 있으나(openapi 결합쿼리 매칭) "
                            "두 번의 자동 브라우저 검증에서 전용 탭 1~10위 노출을 확정하지 못함."
                        )
                        ev_kw["blog"] = ev_blog
                print(
                    f"[playwright_verify] 자동 보정 {corrected}개 · 미확정 {uncertain}개",
                    flush=True,
                )
            except ImportError as e:
                print(f"[playwright_verify] playwright 미설치 (스킵): {e!r}", flush=True)
            except Exception as e:
                print(f"[playwright_verify] 실행 실패 (스킵): {e!r}", flush=True)

    return out, ev_all


def _keyword_channels_for_payload(cfg: dict[str, Any]) -> dict[str, str] | None:
    raw = cfg.get("keywordChannels")
    if not isinstance(raw, dict):
        return None
    clean: dict[str, str] = {}
    for k, v in raw.items():
        ks = str(k).strip()
        vs = str(v).strip().lower()
        if ks and vs in ("cafe", "blog", "all"):
            clean[ks] = vs
    return clean or None


def _keyword_scopes_for_payload(cfg: dict[str, Any]) -> dict[str, str] | None:
    raw = cfg.get("keywordScopes")
    if not isinstance(raw, dict):
        return None
    clean: dict[str, str] = {}
    for k, v in raw.items():
        ks = str(k).strip()
        vs = str(v).strip().lower()
        if ks and vs in ("regional", "national", "other", "all"):
            clean[ks] = vs
    return clean or None


def _load_existing_hospital(hospital_name: str) -> dict[str, Any] | None:
    path = ROOT / "data" / "scoring-data.json"
    if not path.exists():
        return None
    try:
        with open(path, encoding="utf-8") as f:
            root = json.load(f)
    except Exception:
        return None
    target_hn = str(hospital_name or "").strip()
    matches: list[dict[str, Any]] = []
    for record in root.get("months") or []:
        raw_hn = str(record.get("hospitalName") or "").strip()
        belongs = raw_hn == target_hn or (target_hn == "포인트병원" and not raw_hn)
        if belongs:
            matches.append(record)
    if not matches:
        return None

    def _legacy_order(record: dict[str, Any]) -> int:
        match = re.fullmatch(r"(\d{1,2})월", str(record.get("monthLabel") or "").strip())
        return int(match.group(1)) if match else -1

    # 최신 레거시 레코드의 점수를 우선하고, 이전 레코드에만 있던 키워드는 재사용한다.
    sheets: list[dict[str, Any]] = []
    sheets_by_key: dict[str, dict[str, Any]] = {}
    seen_by_key: dict[str, set[str]] = {}
    for record in reversed(sorted(matches, key=_legacy_order)):
        for sheet in record.get("sheets") or []:
            key = str(sheet.get("key") or "").strip()
            if not key:
                continue
            target = sheets_by_key.get(key)
            if target is None:
                target = {
                    "key": key,
                    "title": sheet.get("title"),
                    "header": sheet.get("header"),
                    "rows": [],
                }
                sheets_by_key[key] = target
                seen_by_key[key] = set()
                sheets.append(target)
            for row in sheet.get("rows") or []:
                kw = str((row[2] if len(row) > 2 else "") or "").strip()
                if not kw or kw in seen_by_key[key]:
                    continue
                seen_by_key[key].add(kw)
                target["rows"].append(list(row))
    return {"sheets": sheets}


def _keyword_rows_by_sheet(month: dict[str, Any]) -> dict[str, dict[str, list[Any]]]:
    out: dict[str, dict[str, list[Any]]] = {}
    for s in month.get("sheets") or []:
        sheet_key = str(s.get("key") or "").strip()
        if not sheet_key:
            continue
        rows = out.setdefault(sheet_key, {})
        for row in s.get("rows") or []:
            kw = str((row[2] if len(row) > 2 else "") or "").strip()
            if kw and kw not in rows:
                rows[kw] = list(row)
    return out


def build_month_payload(
    cfg: dict[str, Any],
    ranks_by_device: dict[str, Any],
    volumes: dict[str, dict[str, Any]],
    reused_rows: dict[str, dict[str, list[Any]]] | None = None,
) -> dict[str, Any]:
    def ranks_for_sheet(sheet_key: str) -> dict[str, dict[str, int | None]]:
        device = "mobile" if sheet_key.endswith("-mob") else "pc"
        device_ranks = ranks_by_device.get(device)
        if isinstance(device_ranks, dict):
            return device_ranks
        # 구형 호출 호환: {keyword: {tab: rank}}
        return ranks_by_device  # type: ignore[return-value]

    rbk = cfg.get("rowsBySheetKey") or {}
    titles_override = cfg.get("sheetTitles") or {}
    if rbk and all(k in rbk for k, _ in SHEETS_META):
        sheets = []
        for key, default_title in SHEETS_META:
            ranks = ranks_for_sheet(key)
            title = titles_override.get(key) or default_title
            pairs = rbk[key]
            rows = []
            for i, pair in enumerate(pairs, start=1):
                if isinstance(pair, (list, tuple)):
                    reg, kw = pair[0], pair[1]
                else:
                    reg = pair.get("region")
                    kw = pair.get("keyword") or ""
                reg_s = (str(reg).strip() if reg is not None else "") or ""
                reused_row = (reused_rows or {}).get(key, {}).get(kw)
                if reused_row:
                    row = list(reused_row)
                    if len(row) < 16:
                        row = (row + [None] * 16)[:16]
                    row[0] = i
                    row[1] = reg_s if reg_s else row[1]
                    row[2] = kw
                    rows.append(row)
                    continue
                pts = {tab: table_cell_for_tab(tab, ranks.get(kw, {}).get(tab)) for tab in COL_BY_TAB.keys()}
                v = volumes.get(kw, {})
                rows.append(build_row(i, reg_s, kw, pts, v.get("pc"), v.get("mobile"), v.get("related")))
            sheets.append({"key": key, "title": title, "header": HEADER, "rows": rows})
        out = {
            "sourceFile": cfg.get("sourceFileNote", "배점표_자동(API+WEB).json"),
            "sheets": sheets,
        }
        hn = (cfg.get("hospitalName") or "").strip()
        if hn:
            out["hospitalName"] = hn
        kc = _keyword_channels_for_payload(cfg)
        if kc:
            out["keywordChannels"] = kc
        ks = _keyword_scopes_for_payload(cfg)
        if ks:
            out["keywordScopes"] = ks
        return out

    region = cfg.get("regionDefault", "") or ""
    sheets = []
    for key, title in SHEETS_META:
        ranks = ranks_for_sheet(key)
        rows = []
        for i, kw in enumerate(cfg.get("keywords") or [], start=1):
            reused_row = (reused_rows or {}).get(key, {}).get(kw)
            if reused_row:
                row = list(reused_row)
                if len(row) < 16:
                    row = (row + [None] * 16)[:16]
                row[0] = i
                row[2] = kw
                rows.append(row)
                continue
            pts = {tab: table_cell_for_tab(tab, ranks.get(kw, {}).get(tab)) for tab in COL_BY_TAB.keys()}
            v = volumes.get(kw, {})
            rows.append(build_row(i, region, kw, pts, v.get("pc"), v.get("mobile"), v.get("related")))
        sheets.append({"key": key, "title": title, "header": HEADER, "rows": rows})
    out = {"sourceFile": cfg.get("sourceFileNote", "배점표_자동(API+WEB).json"), "sheets": sheets}
    hn = (cfg.get("hospitalName") or "").strip()
    if hn:
        out["hospitalName"] = hn
    kc = _keyword_channels_for_payload(cfg)
    if kc:
        out["keywordChannels"] = kc
    ks = _keyword_scopes_for_payload(cfg)
    if ks:
        out["keywordScopes"] = ks
    return out


def _write_json_atomic(path: Path, payload: Any, *, compact: bool = False) -> None:
    """임시 파일에 다 쓴 뒤 교체 — 쓰는 도중 프로세스가 죽어도 기존 파일이 잘리지 않는다."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        if compact:
            json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))
        else:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def merge_into_scoring_data(month: dict[str, Any], path: Path | None = None) -> None:
    """동일 병원 항목을 하나로 교체하고 다른 병원 데이터는 유지한다.

    path 를 주면 SCORING_TEMP_OUTPUT 을 무시하고 그 파일에 바로 병합한다 (merge_results.py 용).
    """
    temp_out = os.getenv("SCORING_TEMP_OUTPUT", "").strip()
    if temp_out and path is None:
        _write_json_atomic(Path(temp_out), month)
        print("임시 채점 결과 저장:", temp_out)
        return
    path = path or ROOT / "data" / "scoring-data.json"
    if path.exists():
        with open(path, encoding="utf-8") as f:
            root = json.load(f)
    else:
        root = {"months": []}
    new_hn = str(month.get("hospitalName") or "").strip()
    month.pop("monthLabel", None)

    def _keep_existing(m: dict[str, Any]) -> bool:
        raw_hn = str(m.get("hospitalName") or "").strip()
        belongs = raw_hn == new_hn or (new_hn == "포인트병원" and not raw_hn)
        return not belongs

    months = [m for m in (root.get("months") or []) if _keep_existing(m)]
    months.append(month)
    root["months"] = months
    root["generatedBy"] = "build_month.py(api+web-evidence)"
    _write_json_atomic(path, root)
    print("병합 완료:", path)


def save_evidence(evidence: dict[str, Any], hospital_name: str | None = None, path: Path | None = None) -> None:
    """
    hospital_name 이 있으면 evidence 를 byHospital[hospital_name] 에만 갱신(포인트 공용 evidence 유지).
    없으면 기존처럼 최상위 evidence 전체 교체, byHospital 은 유지.
    path 를 주면 EVIDENCE_TEMP_OUTPUT 을 무시하고 그 파일에 바로 병합한다 (merge_results.py 용).
    파일이 수십 MB 라 들여쓰기 없이 저장한다.
    """
    temp_out = os.getenv("EVIDENCE_TEMP_OUTPUT", "").strip()
    if temp_out and path is None:
        _write_json_atomic(Path(temp_out), {"evidence": evidence, "hospitalName": hospital_name}, compact=True)
        print("임시 근거 저장:", temp_out)
        return
    path = path or ROOT / "data" / "last-run-evidence.json"
    flat: dict[str, Any] = {}
    by_h: dict[str, Any] = {}
    if path.exists() and path.stat().st_size > 0:
        # 기존 파일을 못 읽으면 중단 — 무시하고 저장하면 다른 병원 근거가 전부 지워진다.
        with open(path, encoding="utf-8") as f:
            old = json.load(f)
        flat = dict(old.get("evidence") or {})
        by_h = dict(old.get("byHospital") or {})
    if hospital_name:
        key = str(hospital_name).strip()
        old_h = by_h.get(key)
        merged_h = dict(old_h) if isinstance(old_h, dict) else {}
        merged_h.update(evidence or {})
        by_h[key] = merged_h
    else:
        flat = evidence
    payload = {"generatedAt": time.strftime("%Y-%m-%d %H:%M:%S"), "evidence": flat, "byHospital": by_h}
    _write_json_atomic(path, payload, compact=True)
    print("근거 저장:", path)


def run_scoring_pipeline(cfg: dict[str, Any]) -> None:
    cid = (os.getenv("NAVER_CLIENT_ID") or "").strip()
    csec = (os.getenv("NAVER_CLIENT_SECRET") or "").strip()
    if not cid or not csec:
        raise SystemExit("NAVER_CLIENT_ID / NAVER_CLIENT_SECRET 이 필요합니다.")

    ad_api_key = (os.getenv("NAVER_AD_API_KEY") or "").strip()
    ad_secret = (os.getenv("NAVER_AD_SECRET_KEY") or "").strip()
    ad_customer = (os.getenv("NAVER_AD_CUSTOMER_ID") or "").strip()

    label = (cfg.get("hospitalName") or "배점").strip()
    print(f"네이버 자동 채점 시작 - {label}")
    full_rescore = (os.getenv("SCORING_FULL_RESCORE") or "").strip().lower() in {"1", "true", "yes"}
    existing_month = _load_existing_hospital(label) if label else None
    reused_rows = {} if full_rescore else (_keyword_rows_by_sheet(existing_month) if existing_month else {})
    all_keywords = [str(k).strip() for k in (cfg.get("keywords") or []) if str(k).strip()]
    force_rescore_keywords = {
        str(k).strip() for k in (cfg.get("forceRescoreKeywords") or []) if str(k).strip()
    }
    # forceRescoreKeywords 는 새 점수를 써야 하므로 재사용 대상에서 제외
    # (build_month_payload 는 reused_rows 에 있으면 무조건 기존 row 를 사용함)
    if force_rescore_keywords and reused_rows:
        reused_rows = {
            sheet_key: {
                kw: row for kw, row in rows.items() if kw not in force_rescore_keywords
            }
            for sheet_key, rows in reused_rows.items()
        }
    reused_keywords = {
        kw for rows in reused_rows.values() for kw in rows
    }
    fresh_keywords = (
        all_keywords
        if full_rescore
        else [kw for kw in all_keywords if (kw not in reused_keywords) or (kw in force_rescore_keywords)]
    )
    if full_rescore:
        print(f"전체 재채점 모드: {len(all_keywords)}건")
    else:
        if reused_rows:
            print(f"기존 재사용 키워드: {len(all_keywords) - len(fresh_keywords)}건")
        if force_rescore_keywords:
            print(f"업로드 강제 재채점 키워드: {len(force_rescore_keywords)}건")
        print(f"신규 채점 키워드: {len(fresh_keywords)}건")

    cfg_for_fetch = dict(cfg)
    cfg_for_fetch["keywords"] = fresh_keywords
    pc_ranks, pc_evidence = fetch_keyword_ranks(
        cfg_for_fetch, cid, csec, report_progress=True, device="pc"
    )
    mobile_ranks, mobile_evidence = fetch_keyword_ranks(
        cfg_for_fetch,
        cid,
        csec,
        report_progress=True,
        device="mobile",
        shared_api=(pc_ranks, pc_evidence),
    )
    ranks_by_device = {"pc": pc_ranks, "mobile": mobile_ranks}
    evidence = {
        kw: {
            "pc": pc_evidence.get(kw) or {},
            "mobile": mobile_evidence.get(kw) or {},
        }
        for kw in fresh_keywords
    }
    for kw in fresh_keywords:
        print("-", kw, {"pc": pc_ranks.get(kw), "mobile": mobile_ranks.get(kw)})

    if ad_api_key and ad_secret and ad_customer:
        volumes = fetch_keyword_volumes_searchad(fresh_keywords, ad_api_key, ad_secret, ad_customer)
        print("월간조회수 수집 완료")
    else:
        print("검색광고 API 키 없음 -> 월간조회수는 0")
        volumes = {kw: {"pc": 0, "mobile": 0, "related": kw} for kw in fresh_keywords}

    month = build_month_payload(
        cfg, ranks_by_device, volumes, (None if full_rescore else reused_rows)
    )

    # 배포 안전장치
    # ① 측정 실패 비율: 조회 실패는 0점이 아니라 None('—')으로 남기므로, 실패가 많으면
    #    일치율과 별개로 LOW 처리한다 ("전부 실패했는데 품질 OK" 방지).
    # ② 라이브 재조회 샘플의 시간 일치율.
    # ③ HTTP/API와 브라우저 독립 수집기의 일치율.
    # LOW 면 결과를 저장하지 않고 종료한다 (SCORING_BLOCK_ON_LOW_QUALITY=0 으로 끌 수 있음).
    verify_enabled = (os.getenv("SCORING_VERIFY_SAMPLE") or "1").strip().lower() in {"1", "true", "yes"}
    verify_warn_threshold = float((os.getenv("SCORING_VERIFY_WARN_THRESHOLD") or "0.90").strip() or "0.90")
    verify_low_threshold = float((os.getenv("SCORING_VERIFY_LOW_THRESHOLD") or "0.80").strip() or "0.80")
    verify_size = int((os.getenv("SCORING_VERIFY_SIZE") or "24").strip() or "24")
    max_fail_ratio = float((os.getenv("SCORING_MAX_FAIL_RATIO") or "0.20").strip() or "0.20")
    block_on_low = (os.getenv("SCORING_BLOCK_ON_LOW_QUALITY") or "1").strip().lower() in {"1", "true", "yes"}
    quality_meta: dict[str, Any] | None = None

    measured_cells = 0
    failed_cells = 0
    for device_ranks in ranks_by_device.values():
        for kw in fresh_keywords:
            kw_ranks = device_ranks.get(kw) or {}
            for tab in COL_BY_TAB.keys():
                if tab == "cafe":  # 전역 비활성화 채널
                    continue
                measured_cells += 1
                if kw_ranks.get(tab) is None:
                    failed_cells += 1
    fail_ratio = (failed_cells / measured_cells) if measured_cells else 0.0
    too_many_failures = fail_ratio > max_fail_ratio
    if measured_cells:
        print(
            f"측정 실패 비율: {fail_ratio*100:.1f}% "
            f"({failed_cells}/{measured_cells}칸, 허용 {max_fail_ratio*100:.1f}%)"
        )

    replay_acc: float | None = None
    collector_acc: float | None = None
    sample_size = 0
    total = 0
    verify_pool = all_keywords if full_rescore else fresh_keywords
    if verify_enabled and verify_pool:
        sample = sorted(
            verify_pool,
            key=lambda kw: hashlib.sha256(f"{label}:{kw}".encode("utf-8")).hexdigest(),
        )
        sample = sample[: max(1, min(len(sample), verify_size))]
        sample_size = len(sample)
        cfg_verify = dict(cfg)
        cfg_verify["keywords"] = sample
        replay_pc, replay_pc_ev = fetch_keyword_ranks(
            cfg_verify, cid, csec, device="pc", browser_verify=False
        )
        replay_mobile, _ = fetch_keyword_ranks(
            cfg_verify,
            cid,
            csec,
            device="mobile",
            shared_api=(replay_pc, replay_pc_ev),
            browser_verify=False,
        )
        replay_by_device = {"pc": replay_pc, "mobile": replay_mobile}
        matched = 0
        for device in ("pc", "mobile"):
            for kw in sample:
                for tab in COL_BY_TAB.keys():
                    v1 = table_cell_for_tab(tab, (ranks_by_device[device].get(kw) or {}).get(tab))
                    v2 = table_cell_for_tab(tab, (replay_by_device[device].get(kw) or {}).get(tab))
                    # 한쪽이라도 측정 실패면 비교 불가 — 실패는 위의 실패 비율로 따로 판정한다.
                    if v1 is None or v2 is None:
                        continue
                    total += 1
                    if v1 == v2:
                        matched += 1
        replay_acc = (matched / total) if total else 1.0
        print(
            f"샘플 재조회 일치율: {replay_acc*100:.1f}% "
            f"(경고 {verify_warn_threshold*100:.1f}% / 낮음 {verify_low_threshold*100:.1f}%, 샘플 {sample_size}개)"
        )

    collector_total = 0
    collector_matched = 0
    for kw_devices in evidence.values():
        for device_ev in kw_devices.values():
            if not isinstance(device_ev, dict):
                continue
            for tab_ev in device_ev.values():
                if not isinstance(tab_ev, dict):
                    continue
                audit = tab_ev.get("automationAudit")
                if not isinstance(audit, dict) or not audit.get("agreed"):
                    continue
                collector_total += 1
                if audit.get("primaryAgreed"):
                    collector_matched += 1
    if collector_total:
        collector_acc = collector_matched / collector_total
        print(
            f"독립 수집기 일치율: {collector_acc*100:.1f}% "
            f"({collector_matched}/{collector_total}칸)",
            flush=True,
        )

    if replay_acc is not None or collector_acc is not None or too_many_failures:
        accuracy_values = [x for x in (replay_acc, collector_acc) if x is not None]
        acc = min(accuracy_values) if accuracy_values else 1.0
        if too_many_failures or acc < verify_low_threshold:
            level = "low"
        elif acc < verify_warn_threshold:
            level = "warn"
        else:
            level = "ok"
        quality_meta = {
            "level": level,
            "accuracyPct": round(acc * 100, 1),
            "warnThresholdPct": round(verify_warn_threshold * 100, 1),
            "lowThresholdPct": round(verify_low_threshold * 100, 1),
            "sampleSize": sample_size,
            "checkedItems": total,
            "failedCellPct": round(fail_ratio * 100, 1),
            "temporalAgreementPct": round((replay_acc or 0) * 100, 1) if replay_acc is not None else None,
            "collectorAgreementPct": round((collector_acc or 0) * 100, 1) if collector_acc is not None else None,
            "collectorCheckedItems": collector_total,
            "validationMode": "automated_multi_source",
        }
        print(f"QUALITY_GATE:{level.upper()}:{acc*100:.1f}")
        if level == "low" and block_on_low:
            reason = "측정 실패 과다" if too_many_failures else "자동 교차검증 일치율 미달"
            print(f"품질 게이트 LOW ({reason}) → 채점 결과를 저장하지 않고 종료합니다.", flush=True)
            raise SystemExit(3)

    hn = (cfg.get("hospitalName") or "").strip()
    if hn:
        month["hospitalName"] = hn
    if quality_meta:
        month["quality"] = quality_meta
    merge_into_scoring_data(month)
    save_evidence(evidence, hn or None)


def main() -> None:
    cfg = load_config()
    run_scoring_pipeline(cfg)


if __name__ == "__main__":
    main()
