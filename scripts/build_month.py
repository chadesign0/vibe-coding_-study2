# -*- coding: utf-8 -*-
"""4월 배점표 생성: 네이버 API + 웹 파싱 자동 채점(근거 저장)."""
from __future__ import annotations

import base64
import hashlib
import hmac
import html
import json
import os
import random
import re
import time
from datetime import date, datetime
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
SEARCH_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
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
    블로그 채점 시 인정할 게시 연·월. monthLabel(예: 4월) + scoringYear(없으면 올해).
    blogRequirePostMonth가 false이면 None → 작성일 무시(기존 동작).
    """
    if not cfg.get("blogRequirePostMonth", True):
        return None
    m = re.match(r"^(\d{1,2})월\s*$", str(cfg.get("monthLabel") or "").strip())
    if not m:
        return None
    month = int(m.group(1))
    if not (1 <= month <= 12):
        return None
    env_y = (os.getenv("SCORING_YEAR") or "").strip()
    if env_y.isdigit():
        year = int(env_y)
    else:
        y = cfg.get("scoringYear")
        if y is not None and str(y).strip() != "":
            year = int(y)
        else:
            year = datetime.now().year
    return (year, month)


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
    if rank == 1:
        return 3
    if rank <= 5:
        return 2
    if rank <= 10:
        return 1
    return 0


def table_cell_for_tab(tab: str, raw_rank: int | None) -> int | None:
    """표에 넣을 값. 파워링크만 순위(0=미노출, 1~10), 나머지 탭은 점수(0~3)."""
    if tab == "powerlink":
        if raw_rank is None:
            return None
        return int(raw_rank)
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
        ts = str(int(time.time() * 1000))
        sig = build_searchad_signature(ts, "GET", endpoint, secret_key)
        headers = {"X-Timestamp": ts, "X-API-KEY": api_key, "X-Customer": customer_id, "X-Signature": sig}
        params = {"hintKeywords": kw, "showDetail": 1}
        try:
            r = requests.get("https://api.searchad.naver.com" + endpoint, headers=headers, params=params, timeout=30)
            if r.status_code >= 400:
                result[kw] = {"pc": None, "mobile": None, "related": None}
                continue
            items = r.json().get("keywordList") or []
            picked = next((x for x in items if (x.get("relKeyword") or "") == kw), None) or (items[0] if items else None)
            if not picked:
                result[kw] = {"pc": 0, "mobile": 0, "related": kw}
                continue
            result[kw] = {"pc": parse_count(picked.get("monthlyPcQcCnt")), "mobile": parse_count(picked.get("monthlyMobileQcCnt")), "related": picked.get("relKeyword") or kw}
        except Exception:
            result[kw] = {"pc": None, "mobile": None, "related": None}
    return result


def api_search(client_id: str, client_secret: str, endpoint: str, query: str) -> list[dict[str, Any]] | None:
    url = f"https://openapi.naver.com/v1/search/{endpoint}.json"
    headers = {"X-Naver-Client-Id": client_id, "X-Naver-Client-Secret": client_secret}
    params = {"query": query, "display": 100, "sort": "sim"}
    r = requests.get(url, headers=headers, params=params, timeout=30)
    if r.status_code >= 400:
        return None
    return r.json().get("items") or []


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
    """블로그 API items 상위 10개: 공식 네이버 블로그 ID 일치 또는 작성자·링크에 병원 토큰이 있을 때만 인정."""
    top: list[dict[str, Any]] = []
    matched = 0
    extra: dict[str, Any] = {
        "blogMatchRule": "official_blog_url_or_author_identity",
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
        auth_ok = official_ok or token_ok
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
            "matchAuthorOrBlog": auth_ok,
            "matchTitleOrBodyOnly": bool(body_ok and not auth_ok),
            "bloggername": strip_html(it.get("bloggername", "")),
            "bloggerlink": strip_html(it.get("bloggerlink", "")),
        }
        top.append(row)
        if matched == 0 and in_m and auth_ok:
            matched = i
            extra["matched_text"] = auth_txt[:280]
            extra["matched_postdate"] = it.get("postdate")
            if official_ok:
                extra["matchedVia"] = "official_hospital_blog_url"
            else:
                extra["matchedVia"] = "bloggername_or_bloglink_tokens"
    out: dict[str, Any] = {"top": top, "matched_rank": matched, **extra}
    if matched == 0:
        body_only_in_month = False
        author_out_month = False
        for it in items[:10]:
            auth_txt = blog_author_blog_text_for_match(it)
            full_txt = item_text_for_tab("blog", it)
            official_ok = blog_item_matches_official_naver_blog(it, official_blog_ids)
            token_ok = tokens_match_in_normalized(auth_txt, match_tokens)
            auth_ok = official_ok or token_ok
            body_ok = tokens_match_in_normalized(full_txt, match_tokens)
            in_m = blog_item_in_scoring_month(it, blog_period[0], blog_period[1]) if blog_period else True
            if in_m and body_ok and not auth_ok:
                body_only_in_month = True
            if blog_period and not in_m and auth_ok:
                author_out_month = True
        if body_only_in_month:
            out["blogNote"] = (
                "상위 노출 중 제목·본문에만 병원명이 있고, 공식 블로그(hospitalBlogBases) 및 작성자·블로그 링크 기준과 맞지 않아 블로그 0점 처리."
            )
        elif author_out_month:
            out["blogNote"] = (
                "공식 블로그 또는 작성자/블로그 정보 기준은 맞으나 네이버 표기 작성일(postdate)이 "
                f"{blog_period[0]}년 {blog_period[1]}월이 아님."
            )
        elif blog_period:
            out["blogNote"] = (
                f"배점 월({blog_period[0]}년 {blog_period[1]}월) 내 작성이면서 공식 블로그·작성자/블로그 기준에 맞는 글이 없음."
            )
    return (matched if matched else 0), out


def find_rank_by_api_tab(
    tab: str,
    query: str,
    match_tokens: list[str],
    client_id: str,
    client_secret: str,
    blog_period: tuple[int, int] | None = None,
    official_blog_ids: frozenset[str] = frozenset(),
    official_cafe_ids: frozenset[str] = frozenset(),
) -> tuple[int | None, dict[str, Any]]:
    endpoint = TAB_ENDPOINT.get(tab)
    if not endpoint:
        return None, {"reason": "unsupported endpoint"}
    items = api_search(client_id, client_secret, endpoint, query)
    if items is None:
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
    return (matched if matched else 0), {"top": top, "matched_rank": matched, **extra}


def fetch_search_page(query: str, where: str | None = None) -> str | None:
    base = "https://search.naver.com/search.naver?query=" + quote_plus(query)
    urls = [base]
    if where:
        urls.insert(0, base + "&where=" + quote_plus(where))
    for _ in range(3):
        for url in urls:
            try:
                r = requests.get(url, headers=SEARCH_HEADERS, timeout=30)
                if r.status_code >= 400:
                    continue
                return decode_response_text(r)
            except Exception:
                continue
        time.sleep(0.4)
    return None


def fetch_integrated_search_page(query: str) -> str | None:
    urls = [
        "https://search.naver.com/search.naver?where=nexearch&sm=tab_jum&ssc=tab.nx.all&query=" + quote_plus(query),
        "https://search.naver.com/search.naver?query=" + quote_plus(query),
    ]
    for _ in range(3):
        for url in urls:
            try:
                r = requests.get(url, headers=SEARCH_HEADERS, timeout=30)
                if r.status_code >= 400:
                    continue
                return decode_response_text(r)
            except Exception:
                continue
        time.sleep(0.4)
    return None


def fetch_powerlink_more_page(query: str) -> str | None:
    """
    파워링크는 통합검색 메인 블록이 아닌 '더보기(광고 전체)' 기준으로 순위를 산정한다.
    """
    url = "https://ad.search.naver.com/search.naver?where=ad&query=" + quote_plus(query)
    try:
        r = requests.get(url, headers=SEARCH_HEADERS, timeout=30)
        if r.status_code >= 400:
            return None
        enc = (r.apparent_encoding or r.encoding or "utf-8").strip()
        try:
            return r.content.decode(enc, errors="ignore")
        except Exception:
            return r.text
    except Exception:
        return None


def extract_candidates_powerlink(ht: str) -> list[str]:
    soup = BeautifulSoup(ht, "html.parser")
    root = soup.select_one("div[id^='pcPowerLink_']")
    if not root:
        return []
    vals = []
    for n in root.select("li, .ad_dsc, .url_area, .desc"):
        t = strip_html(n.get_text(" ", strip=True))
        if t:
            vals.append(t)
    return vals[:20]


def extract_candidates_powerlink_more(ht: str) -> list[str]:
    """
    ad.search.naver.com(where=ad) 페이지의 광고 목록 순서 그대로 후보를 추출한다.
    """
    soup = BeautifulSoup(ht, "html.parser")
    vals: list[str] = []
    for sel in ("li.lst", "ul.lst_type > li", "div.ad_list li", "ol > li"):
        nodes = soup.select(sel)
        if not nodes:
            continue
        for n in nodes:
            t = strip_html(n.get_text(" ", strip=True))
            if t and len(t) > 3:
                vals.append(t)
        if vals:
            break
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

    # 2) 렌더 데이터 fallback
    marker = "review/prs_template_v2_review_blog_rra_desk.ts"
    idx = ht.find(marker)
    if idx < 0:
        marker = "review/prs_template_v2_review_blog_tab_desk.ts"
        idx = ht.find(marker)
    if idx < 0:
        marker = "review/prs_template_v2_review_ugc_single_intention_desk.ts"
        idx = ht.find(marker)
    if idx < 0:
        marker = "review/prs_template_v2_review_ugc_single_intention_mob.ts"
        idx = ht.find(marker)
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


def find_rank_by_web_tab(
    tab: str,
    query: str,
    match_tokens: list[str],
    blog_period: tuple[int, int] | None = None,
    official_blog_ids: frozenset[str] = frozenset(),
) -> tuple[int | None, dict[str, Any]]:
    if tab == "powerlink":
        ht_more = fetch_powerlink_more_page(query)
        if ht_more:
            cands = extract_candidates_powerlink_more(ht_more)
            rank = find_rank_in_candidates(cands, match_tokens)
            top = [{"rank": i + 1, "text": t[:220]} for i, t in enumerate(cands[:10])]
            return rank, {"top": top, "matched_rank": rank, "basis": "powerlink_more"}
        # 더보기 페이지 실패 시 기존 통합검색 블록 파서로 폴백
        ht = fetch_search_page(query)
        if not ht:
            return None, {"reason": "http_error"}
        cands = extract_candidates_powerlink(ht)
        rank = find_rank_in_candidates(cands, match_tokens)
        top = [{"rank": i + 1, "text": t[:220]} for i, t in enumerate(cands[:10])]
        return rank, {"top": top, "matched_rank": rank, "basis": "powerlink_main_fallback"}
    if tab == "web":
        # 통합검색 웹 블록은 응답 변동이 커서 여러 번 재시도한다.
        last_cands: list[str] = []
        for _ in range(3):
            ht = fetch_integrated_search_page(query)
            if not ht:
                time.sleep(0.4)
                continue
            cands = extract_candidates_web_from_integrated(ht)
            if cands:
                rank = find_rank_in_candidates(cands, match_tokens)
                top = [{"rank": i + 1, "text": t[:220]} for i, t in enumerate(cands[:10])]
                return rank, {"top": top, "matched_rank": rank, "basis": "web_integrated"}
            last_cands = cands
            time.sleep(0.4)

        # 폴백: 통합검색 파싱 실패 시 블로그 API를 보조 근거로 사용(0 오검출 완화)
        items = api_search(
            (os.getenv("NAVER_CLIENT_ID") or "").strip(),
            (os.getenv("NAVER_CLIENT_SECRET") or "").strip(),
            TAB_ENDPOINT["blog"],
            query,
        )
        if items:
            rank, ev = analyze_blog_search_items(items, match_tokens, blog_period, official_blog_ids)
            return rank, {**ev, "basis": "web_blog_fallback"}

        return 0, {"top": [{"rank": i + 1, "text": t[:220]} for i, t in enumerate(last_cands[:10])], "matched_rank": 0, "reason": "parse_empty"}

    ht = fetch_search_page(query, where=("video" if tab == "video" else None))
    if not ht:
        return None, {"reason": "http_error"}
    if tab == "bizsite":
        cands = extract_candidates_bizsite(ht)
    elif tab == "video":
        cands = extract_candidates_video(ht)
    else:
        return None, {"reason": "unsupported_web_tab"}
    rank = find_rank_in_candidates(cands, match_tokens)
    top = [{"rank": i + 1, "text": t[:220]} for i, t in enumerate(cands[:10])]
    return rank, {"top": top, "matched_rank": rank}


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


def fetch_keyword_ranks(cfg: dict[str, Any], cid: str, csec: str) -> tuple[dict[str, dict[str, int | None]], dict[str, Any]]:
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
    for kw in cfg.get("keywords") or []:
        out[kw], ev_all[kw] = {}, {}
        for tab in COL_BY_TAB.keys():
            if tab == "cafe":
                out[kw][tab] = 0
                ev_all[kw][tab] = {
                    "source": "disabled",
                    "reason": "cafe_removed",
                    "matched_rank": 0,
                    "top": [],
                    "note": "카페 채점 전역 비활성화",
                }
                continue
            if tab in manual_by_tab.get(kw, {}):
                r = int(manual_by_tab[kw][tab]); out[kw][tab] = r; ev_all[kw][tab] = {"source": "manual", "rank": r}; continue
            if tab == "blog" and kw in legacy_blog_manual:
                r = int(legacy_blog_manual[kw]); out[kw][tab] = r; ev_all[kw][tab] = {"source": "manual_legacy_blog", "rank": r}; continue
            if tab in ("powerlink", "bizsite", "video", "web"):
                r, ev = find_rank_by_web_tab(tab, kw, match_tokens, blog_period, official_blog_ids)
                out[kw][tab] = 0 if r is None else r
                ev_all[kw][tab] = {"source": "web", **ev}
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
                )
                out[kw][tab] = 0 if r is None else r
                ev_all[kw][tab] = {"source": "api", **ev}
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


def _load_existing_month(month_label: str, hospital_name: str) -> dict[str, Any] | None:
    path = ROOT / "data" / "scoring-data.json"
    if not path.exists():
        return None
    try:
        with open(path, encoding="utf-8") as f:
            root = json.load(f)
    except Exception:
        return None
    target_lab = str(month_label or "").strip()
    target_hn = str(hospital_name or "").strip()
    for m in root.get("months") or []:
        if str(m.get("monthLabel") or "").strip() != target_lab:
            continue
        if str(m.get("hospitalName") or "").strip() != target_hn:
            continue
        return m
    return None


def _keyword_row_index_from_month(month: dict[str, Any]) -> dict[str, list[Any]]:
    out: dict[str, list[Any]] = {}
    for s in month.get("sheets") or []:
        for row in s.get("rows") or []:
            kw = str((row[2] if len(row) > 2 else "") or "").strip()
            if kw and kw not in out:
                out[kw] = list(row)
    return out


def build_month_payload(
    cfg: dict[str, Any],
    ranks: dict[str, dict[str, int | None]],
    volumes: dict[str, dict[str, Any]],
    reused_rows: dict[str, list[Any]] | None = None,
) -> dict[str, Any]:
    rbk = cfg.get("rowsBySheetKey") or {}
    titles_override = cfg.get("sheetTitles") or {}
    if rbk and all(k in rbk for k, _ in SHEETS_META):
        sheets = []
        for key, default_title in SHEETS_META:
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
                if reused_rows and kw in reused_rows:
                    row = list(reused_rows[kw])
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
            "monthLabel": cfg.get("monthLabel", "4월"),
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
    rows = []
    for i, kw in enumerate(cfg.get("keywords") or [], start=1):
        if reused_rows and kw in reused_rows:
            row = list(reused_rows[kw])
            if len(row) < 16:
                row = (row + [None] * 16)[:16]
            row[0] = i
            row[2] = kw
            rows.append(row)
            continue
        pts = {tab: table_cell_for_tab(tab, ranks.get(kw, {}).get(tab)) for tab in COL_BY_TAB.keys()}
        v = volumes.get(kw, {})
        rows.append(build_row(i, region, kw, pts, v.get("pc"), v.get("mobile"), v.get("related")))
    sheets = [{"key": key, "title": title, "header": HEADER, "rows": rows} for key, title in SHEETS_META]
    out = {"sourceFile": cfg.get("sourceFileNote", "4월_배점표_자동(API+WEB).json"), "monthLabel": cfg.get("monthLabel", "4월"), "sheets": sheets}
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


def _month_order_label(label: str) -> int:
    m = re.match(r"^(\d{1,2})월\s*$", (label or "").strip())
    return int(m.group(1)) if m else 99


def _month_identity(m: dict[str, Any]) -> tuple[str, str]:
    """병합 키: (monthLabel, hospitalName).

    - 같은 달이라도 병원명이 다르면 서로 다른 슬롯 → 데이터가 겹치거나 덮어쓰이지 않음.
    - hospitalName 이 비어 있으면 포인트병원 레거시 슬롯(다른 병원과 공존 가능).
    """
    lab = str(m.get("monthLabel") or "").strip()
    h = str(m.get("hospitalName") or "").strip()
    return lab, h


def merge_into_scoring_data(month: dict[str, Any]) -> None:
    """동일 (monthLabel, hospitalName) 항목만 교체·추가. 타 병원·타 슬롯은 유지."""
    temp_out = os.getenv("SCORING_TEMP_OUTPUT", "").strip()
    if temp_out:
        with open(temp_out, "w", encoding="utf-8") as f:
            json.dump(month, f, ensure_ascii=False, indent=2)
        print("임시 채점 결과 저장:", temp_out)
        return
    path = ROOT / "data" / "scoring-data.json"
    with open(path, encoding="utf-8") as f:
        root = json.load(f)
    new_id = _month_identity(month)
    new_lab = str(month.get("monthLabel") or "").strip()
    new_hn = str(month.get("hospitalName") or "").strip()

    def _keep_existing(m: dict[str, Any]) -> bool:
        if _month_identity(m) == new_id:
            return False
        # 명시 포인트병원 블록을 저장할 때, 같은 달·hospitalName 없는 레거시(포인트 전용)는 중복이라 제거
        if new_hn == "포인트병원" and new_lab:
            ml = str(m.get("monthLabel") or "").strip()
            raw = m.get("hospitalName")
            if ml == new_lab and (raw is None or str(raw).strip() == ""):
                return False
        return True

    months = [m for m in (root.get("months") or []) if _keep_existing(m)]
    months.append(month)
    months.sort(key=lambda m: _month_order_label(str(m.get("monthLabel") or "")))
    root["months"] = months
    root["generatedBy"] = "build_april_month.py(api+web-evidence)"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(root, f, ensure_ascii=False, indent=2)
    print("병합 완료:", path)


def save_evidence(evidence: dict[str, Any], hospital_name: str | None = None) -> None:
    """
    hospital_name 이 있으면 evidence 를 byHospital[hospital_name] 에만 갱신(포인트 공용 evidence 유지).
    없으면 기존처럼 최상위 evidence 전체 교체, byHospital 은 유지.
    """
    temp_out = os.getenv("EVIDENCE_TEMP_OUTPUT", "").strip()
    if temp_out:
        with open(temp_out, "w", encoding="utf-8") as f:
            json.dump({"evidence": evidence, "hospitalName": hospital_name}, f, ensure_ascii=False, indent=2)
        print("임시 근거 저장:", temp_out)
        return
    path = ROOT / "data" / "last-run-evidence.json"
    flat: dict[str, Any] = {}
    by_h: dict[str, Any] = {}
    if path.exists():
        try:
            with open(path, encoding="utf-8") as f:
                old = json.load(f)
            flat = dict(old.get("evidence") or {})
            by_h = dict(old.get("byHospital") or {})
        except Exception:
            pass
    if hospital_name:
        key = str(hospital_name).strip()
        old_h = by_h.get(key)
        merged_h = dict(old_h) if isinstance(old_h, dict) else {}
        merged_h.update(evidence or {})
        by_h[key] = merged_h
    else:
        flat = evidence
    payload = {"generatedAt": time.strftime("%Y-%m-%d %H:%M:%S"), "evidence": flat, "byHospital": by_h}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
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
    month_label = str(cfg.get("monthLabel") or "").strip()
    existing_month = _load_existing_month(month_label, label) if month_label and label else None
    reused_rows = {} if full_rescore else (_keyword_row_index_from_month(existing_month) if existing_month else {})
    all_keywords = [str(k).strip() for k in (cfg.get("keywords") or []) if str(k).strip()]
    force_rescore_keywords = {
        str(k).strip() for k in (cfg.get("forceRescoreKeywords") or []) if str(k).strip()
    }
    # forceRescoreKeywords 는 새 점수를 써야 하므로 재사용 대상에서 제외
    # (build_month_payload 는 reused_rows 에 있으면 무조건 기존 row 를 사용함)
    if force_rescore_keywords and reused_rows:
        reused_rows = {
            kw: row for kw, row in reused_rows.items() if kw not in force_rescore_keywords
        }
    fresh_keywords = (
        all_keywords
        if full_rescore
        else [kw for kw in all_keywords if (kw not in reused_rows) or (kw in force_rescore_keywords)]
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
    ranks, evidence = fetch_keyword_ranks(cfg_for_fetch, cid, csec)
    for kw, tabs in ranks.items():
        print("-", kw, tabs)

    if ad_api_key and ad_secret and ad_customer:
        volumes = fetch_keyword_volumes_searchad(fresh_keywords, ad_api_key, ad_secret, ad_customer)
        print("월간조회수 수집 완료")
    else:
        print("검색광고 API 키 없음 -> 월간조회수는 0")
        volumes = {kw: {"pc": 0, "mobile": 0, "related": kw} for kw in fresh_keywords}

    month = build_month_payload(cfg, ranks, volumes, (None if full_rescore else reused_rows))

    # 배포 안전장치: 라이브 재조회 샘플과의 일치율이 임계치 미만이면 실패 처리
    verify_enabled = (os.getenv("SCORING_VERIFY_SAMPLE") or "1").strip().lower() in {"1", "true", "yes"}
    verify_warn_threshold = float((os.getenv("SCORING_VERIFY_WARN_THRESHOLD") or "0.90").strip() or "0.90")
    verify_low_threshold = float((os.getenv("SCORING_VERIFY_LOW_THRESHOLD") or "0.80").strip() or "0.80")
    verify_size = int((os.getenv("SCORING_VERIFY_SIZE") or "24").strip() or "24")
    quality_meta: dict[str, Any] | None = None
    verify_pool = all_keywords if full_rescore else fresh_keywords
    if verify_enabled and verify_pool:
        sample = verify_pool[:]
        random.shuffle(sample)
        sample = sample[: max(1, min(len(sample), verify_size))]
        cfg_verify = dict(cfg)
        cfg_verify["keywords"] = sample
        replay_ranks, _ = fetch_keyword_ranks(cfg_verify, cid, csec)
        matched = 0
        total = 0
        for kw in sample:
            for tab in COL_BY_TAB.keys():
                v1 = table_cell_for_tab(tab, ranks.get(kw, {}).get(tab))
                v2 = table_cell_for_tab(tab, replay_ranks.get(kw, {}).get(tab))
                if v1 is None and v2 is None:
                    continue
                total += 1
                if v1 == v2:
                    matched += 1
        replay_acc = (matched / total) if total else 1.0
        print(
            f"샘플 재조회 일치율: {replay_acc*100:.1f}% "
            f"(경고 {verify_warn_threshold*100:.1f}% / 낮음 {verify_low_threshold*100:.1f}%, 샘플 {len(sample)}개)"
        )
        if replay_acc < verify_low_threshold:
            quality_meta = {
                "level": "low",
                "accuracyPct": round(replay_acc * 100, 1),
                "warnThresholdPct": round(verify_warn_threshold * 100, 1),
                "lowThresholdPct": round(verify_low_threshold * 100, 1),
                "sampleSize": len(sample),
                "checkedItems": total,
            }
            print(f"QUALITY_GATE:LOW:{replay_acc*100:.1f}")
        elif replay_acc < verify_warn_threshold:
            quality_meta = {
                "level": "warn",
                "accuracyPct": round(replay_acc * 100, 1),
                "warnThresholdPct": round(verify_warn_threshold * 100, 1),
                "lowThresholdPct": round(verify_low_threshold * 100, 1),
                "sampleSize": len(sample),
                "checkedItems": total,
            }
            print(f"QUALITY_GATE:WARN:{replay_acc*100:.1f}")
        else:
            quality_meta = {
                "level": "ok",
                "accuracyPct": round(replay_acc * 100, 1),
                "warnThresholdPct": round(verify_warn_threshold * 100, 1),
                "lowThresholdPct": round(verify_low_threshold * 100, 1),
                "sampleSize": len(sample),
                "checkedItems": total,
            }
            print(f"QUALITY_GATE:OK:{replay_acc*100:.1f}")

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
