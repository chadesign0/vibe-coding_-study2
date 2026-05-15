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


async def _extract_blog_tab_official_rank(
    page,
    keyword: str,
    official_blog_ids: frozenset[str],
) -> tuple[int, dict[str, Any]]:
    """블로그 전용 탭(where=blog)을 풀 렌더링해서 공식 블로그 ID 노출 rank 추출.

    DOM의 organic 카드만 카운트 — 광고/추천 영역은 제외.
    1~10위 안에 공식 블로그 ID가 들어오면 그 rank 반환. 없으면 0.

    반환: (rank, evidence). evidence에는 source URL / 추출된 카드 수 / 매칭된 blog id 등 디버깅 정보.
    """
    ev: dict[str, Any] = {"source": "blog_tab_playwright", "where": "search.naver.com?where=blog"}
    if not official_blog_ids:
        return 0, ev
    url = "https://search.naver.com/search.naver?where=blog&query=" + quote_plus(keyword)
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=_GOTO_TIMEOUT)
    except Exception as e:
        ev["error"] = f"goto_failed: {e!r}"
        return 0, ev
    await _scroll_to_bottom(page)
    try:
        # organic 블로그 카드의 blog.naver.com/{bid}/{postId} 링크를 등장 순서대로 수집.
        # `data-url`/`href`에서 추출하고 광고 영역(data-cr-area에 ad 표식)은 제외.
        cards = await page.evaluate(
            """
            () => {
              const out = [];
              const seen = new Set();
              // 광고/추천 영역에 속한 노드는 스킵
              function inAd(el) {
                let cur = el;
                while (cur && cur !== document.body) {
                  const area = cur.getAttribute && cur.getAttribute('data-cr-area');
                  const cls = (cur.className || '').toString();
                  if (area && /^a[0-9a-z]*\\*?$/i.test(area) === false && /^(adi|ad_|ads_|spi_ad)/i.test(area)) return true;
                  if (/spw_ad|search_ad|ad_section|powerlink/i.test(cls)) return true;
                  cur = cur.parentElement;
                }
                return false;
              }
              const re = /^https?:\\/\\/(?:m\\.)?blog\\.naver\\.com\\/([a-zA-Z0-9_\\-]+)\\/(\\d+)/;
              const candidates = document.querySelectorAll('[data-url], a[href]');
              for (const el of candidates) {
                const raw = el.getAttribute('data-url') || el.getAttribute('href') || '';
                const m = raw.match(re);
                if (!m) continue;
                if (inAd(el)) continue;
                const key = m[1] + '/' + m[2];
                if (seen.has(key)) continue;
                seen.add(key);
                out.push({ bid: m[1], postId: m[2] });
                if (out.length >= 30) break;
              }
              return out;
            }
            """
        )
    except Exception as e:
        ev["error"] = f"evaluate_failed: {e!r}"
        return 0, ev
    ev["organicCardsExtracted"] = len(cards or [])
    rank = 0
    matched_bid = None
    matched_post = None
    for i, c in enumerate(cards or [], start=1):
        if i > 10:
            break
        if c.get("bid") in official_blog_ids:
            rank = i
            matched_bid = c.get("bid")
            matched_post = c.get("postId")
            break
    if rank:
        ev["matched_rank"] = rank
        ev["matched_blog_id"] = matched_bid
        ev["matched_post_id"] = matched_post
    return rank, ev


async def _verify_one(
    page,
    keyword: str,
    match_tokens: list[str],
    official_blog_ids: frozenset[str],
    *,
    do_blog_tab: bool = False,
) -> dict[str, dict[str, Any]]:
    """단일 키워드 검증 — 매칭된 channel별 {rank, source, evidence} 반환.

    do_blog_tab=True 이면 통합검색 외에 블로그 전용탭(where=blog)도 확인.
    반환 schema: {channel: {"rank": int, "source": str, "evidence": dict}}
    """
    out: dict[str, dict[str, Any]] = {}
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
        out["web"] = {"rank": rank, "source": "integrated_search_dom"}

    # 2. MAP — drt 메타 8건
    for i, (_pid, name) in enumerate(_DRT_META_RE.findall(html)[:8], start=1):
        if any(t in normalize_text(name) for t in match_tokens):
            out["map"] = {"rank": i, "source": "integrated_search_dom"}
            break

    # 3. POWERLINK 광고 영역 (통합검색 메인)
    pw_rank = _match_first(extract_candidates_powerlink(html), match_tokens)
    if pw_rank > 0:
        out["powerlink"] = {"rank": pw_rank, "source": "integrated_search_dom"}

    # 4. BIZSITE 광고 영역
    bz_rank = _match_first(extract_candidates_bizsite(html), match_tokens)
    if bz_rank > 0:
        out["bizsite"] = {"rank": bz_rank, "source": "integrated_search_dom"}

    # 5. VIDEO 캐러셀
    items = extract_video_items_with_dates(html)
    for i, (txt, _date) in enumerate(items[:10], start=1):
        if any(t in normalize_text(txt) for t in match_tokens):
            out["video"] = {"rank": i, "source": "integrated_search_dom"}
            break

    # 6. BLOG — 통합검색 페이지에서 공식 블로그 ID 등장 (보수적 1위 부여)
    if official_blog_ids:
        for bid in official_blog_ids:
            if f"blog.naver.com/{bid}" in html:
                out["blog"] = {"rank": 1, "source": "integrated_search_dom"}
                break

    # 7. BLOG (강화) — 블로그 전용 탭 풀 렌더링에서 1~10위 안 공식 블로그 ID 매칭
    # 이미 통합검색에서 잡혔어도 더 정확한 rank 산정을 위해 시도
    if do_blog_tab and official_blog_ids:
        bt_rank, bt_ev = await _extract_blog_tab_official_rank(page, keyword, official_blog_ids)
        if bt_rank > 0:
            # 블로그탭 rank를 우선 (사용자 view에 더 가까운 source)
            out["blog"] = {"rank": bt_rank, "source": "blog_tab_dom", "evidence": bt_ev}
        else:
            # 못 찾았으면 evidence만 — 점수 부여 X
            existing = out.get("blog")
            if existing:
                existing["blog_tab_check"] = bt_ev
            else:
                # 통합검색에서도 못 찾고 블로그탭에서도 못 찾았으면 결과에 흔적 남기지 않음
                pass

    return out


async def _verify_all(
    keywords: list[str],
    match_tokens: list[str],
    official_blog_ids: frozenset[str],
    blog_tab_keywords: frozenset[str] = frozenset(),
) -> dict[str, dict[str, dict[str, Any]]]:
    """전체 0점 keyword 순차 처리. 동시성은 throttle 위험으로 1.

    blog_tab_keywords: 이 집합에 속한 키워드는 블로그 전용탭(where=blog) 풀 렌더링도 추가 실행.
      (PostSearchList/openapi 사전 필터를 통과한 키워드만 비싼 블로그탭 verify를 돌림)
    """
    from playwright.async_api import async_playwright

    results: dict[str, dict[str, dict[str, Any]]] = {}
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
                do_blog = kw in blog_tab_keywords
                tag = " +blogTab" if do_blog else ""
                print(f"  [pw {i}/{total}]{tag} {kw}", flush=True)
                try:
                    results[kw] = await _verify_one(
                        page, kw, match_tokens, official_blog_ids, do_blog_tab=do_blog
                    )
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
    blog_tab_keywords: frozenset[str] = frozenset(),
) -> dict[str, dict[str, dict[str, Any]]]:
    """동기 진입점 — build_month.py에서 호출.

    keywords: 1차 채점에서 powerlink/bizsite/map/blog/news/video/web 모두 0점인 키워드.
    blog_tab_keywords: 그 중 블로그 전용탭(where=blog) 추가 verify 대상 (사전 필터 통과).
    반환: {keyword: {channel: {rank, source, evidence?}}}
    """
    if not keywords:
        return {}
    try:
        return asyncio.run(
            _verify_all(keywords, match_tokens, official_blog_ids, blog_tab_keywords)
        )
    except Exception as e:
        print(f"[playwright_verify] 전체 실패 (무시): {e!r}", flush=True)
        return {}
