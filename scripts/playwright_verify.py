# -*- coding: utf-8 -*-
"""Playwright 보조 채점: 1차 채점에서 전 채널 0점인 키워드만 헤드리스 브라우저로 재검증.

NAVER 공식 API + HTTP fetch로는 잡히지 않는 경우 — JS 렌더링/lazy-load 후의 통합검색 페이지 DOM에서
영역별로 hospital 매칭을 다시 시도. 영역→채점 컬럼 매핑은 스코프상 6개 채널만 다룬다.

매핑(현 정책):
- 통합검색 organic 외부 사이트 → web
- 통합검색 plat 카드(drt 메타)        → map
- 통합검색 파워링크 광고 영역         → powerlink
- 통합검색 비즈사이트 광고 영역       → bizsite
- 통합검색 동영상 캐러셀              → video
- 통합검색 페이지 내 공식 블로그 노출  → blog (rank 산정 불가 시 1로 보수적 부여)

blog/news/cafe는 NAVER 공식 검색 API 영역이라 Playwright 검증 대상 아님.
인플루언서 영역은 정책상 무시.
"""
from __future__ import annotations

import asyncio
import re
from typing import Any
from urllib.parse import quote_plus

from build_month import (
    _DRT_META_RE,
    _find_web_rank_by_url,
    _find_web_rank_from_render_json,
    extract_candidates_bizsite,
    extract_candidates_powerlink,
    extract_video_items_with_dates,
    normalize_text,
)

_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

# 페이지당 처리 대기시간 (ms)
_GOTO_TIMEOUT = 30000
_AFTER_LOAD_WAIT = 1500
_BETWEEN_KEYWORD_WAIT = 2000


async def _scroll_to_bottom(page) -> None:
    """lazy-load 영역 트리거 — 5,000px까지 또는 페이지 끝까지 스크롤."""
    try:
        await page.evaluate(
            """
            new Promise((resolve) => {
                let total = 0; const step = 600;
                const t = setInterval(() => {
                    window.scrollBy(0, step);
                    total += step;
                    const atBottom = (window.innerHeight + window.scrollY) >= document.body.scrollHeight - 50;
                    if (total >= 5000 || atBottom) {
                        clearInterval(t);
                        resolve();
                    }
                }, 250);
            });
            """
        )
        await page.wait_for_timeout(_AFTER_LOAD_WAIT)
    except Exception:
        pass


def _match_first(cands: list[str], match_tokens: list[str]) -> int:
    """후보 리스트에서 처음으로 매칭 토큰이 등장하는 1-based rank. 없으면 0."""
    for i, txt in enumerate(cands[:10], start=1):
        if any(t in normalize_text(txt) for t in match_tokens):
            return i
    return 0


async def _verify_one(
    page,
    keyword: str,
    match_tokens: list[str],
    official_blog_ids: frozenset[str],
) -> dict[str, int]:
    """단일 키워드 검증 — 매칭된 channel만 dict에 담아 반환."""
    out: dict[str, int] = {}
    url = (
        "https://search.naver.com/search.naver?where=nexearch&sm=tab_jum"
        "&ssc=tab.nx.all&query=" + quote_plus(keyword)
    )
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=_GOTO_TIMEOUT)
    except Exception:
        return out

    await _scroll_to_bottom(page)
    try:
        html = await page.content()
    except Exception:
        return out

    # 1. WEB — 외부 organic 결과
    rank, _ev = _find_web_rank_by_url(html, match_tokens, official_blog_ids)
    if rank == 0:
        rank, _ev = _find_web_rank_from_render_json(html, match_tokens, official_blog_ids)
    if rank > 0:
        out["web"] = rank

    # 2. MAP — drt 메타 8건
    for i, (_pid, name) in enumerate(_DRT_META_RE.findall(html)[:8], start=1):
        if any(t in normalize_text(name) for t in match_tokens):
            out["map"] = i
            break

    # 3. POWERLINK 광고 영역 (통합검색 메인)
    pw_rank = _match_first(extract_candidates_powerlink(html), match_tokens)
    if pw_rank > 0:
        out["powerlink"] = pw_rank

    # 4. BIZSITE 광고 영역
    bz_rank = _match_first(extract_candidates_bizsite(html), match_tokens)
    if bz_rank > 0:
        out["bizsite"] = bz_rank

    # 5. VIDEO 캐러셀
    items = extract_video_items_with_dates(html)
    for i, (txt, _date) in enumerate(items[:10], start=1):
        if any(t in normalize_text(txt) for t in match_tokens):
            out["video"] = i
            break

    # 6. BLOG — 통합검색 페이지에 공식 네이버 블로그 ID URL 등장 여부
    # 통합검색 페이지에서는 정확한 rank 산정이 어려워 등장 시 1위로 보수적 처리.
    # 공식 ID 매칭이 false positive를 발생시키지 않으므로 안전.
    if official_blog_ids:
        for bid in official_blog_ids:
            if f"blog.naver.com/{bid}" in html:
                out["blog"] = 1
                break

    return out


async def _verify_all(
    keywords: list[str],
    match_tokens: list[str],
    official_blog_ids: frozenset[str],
) -> dict[str, dict[str, int]]:
    """전체 0점 keyword 순차 처리. 동시성은 throttle 위험으로 1."""
    from playwright.async_api import async_playwright

    results: dict[str, dict[str, int]] = {}
    total = len(keywords)
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        try:
            ctx = await browser.new_context(
                user_agent=_UA,
                viewport={"width": 1366, "height": 800},
                locale="ko-KR",
            )
            page = await ctx.new_page()
            for i, kw in enumerate(keywords, 1):
                print(f"  [pw {i}/{total}] {kw}", flush=True)
                try:
                    results[kw] = await _verify_one(page, kw, match_tokens, official_blog_ids)
                except Exception as e:
                    print(f"    예외: {e!r}", flush=True)
                    results[kw] = {}
                await page.wait_for_timeout(_BETWEEN_KEYWORD_WAIT)
        finally:
            await browser.close()
    return results


def verify_zero_keywords(
    keywords: list[str],
    match_tokens: list[str],
    official_blog_ids: frozenset[str],
) -> dict[str, dict[str, int]]:
    """동기 진입점 — build_month.py에서 호출.

    keywords: 1차 채점에서 powerlink/bizsite/map/blog/news/video/web 모두 0점인 키워드.
    반환: {keyword: {channel: rank}} — 매칭된 channel만 포함.
    """
    if not keywords:
        return {}
    try:
        return asyncio.run(_verify_all(keywords, match_tokens, official_blog_ids))
    except Exception as e:
        print(f"[playwright_verify] 전체 실패 (무시): {e!r}", flush=True)
        return {}
