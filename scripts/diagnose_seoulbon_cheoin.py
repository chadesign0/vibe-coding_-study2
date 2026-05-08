# -*- coding: utf-8 -*-
"""서울본정형외과 '처인구재활치료' 키워드 진단 — video + map 매칭 추적.

사용자 보고: 동영상 탭 1위에 서울본정형외과 글이 노출되는데 채점 동영상=0.
"""
import io
import os
import re
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, str(Path(__file__).parent))

from dotenv import load_dotenv
load_dotenv()

from build_month import (
    api_search,
    build_match_tokens,
    extract_video_items_with_dates,
    fetch_search_page,
    fetch_integrated_search_page,
    find_rank_by_api_tab,
    find_rank_by_web_tab,
    item_text_for_tab,
    normalize_text,
    strip_html,
    _DRT_META_RE,
)

KEYWORD = '처인구재활치료'
NAMES = ['서울본정형외과']
DOMAINS = ['https://yonginseoulbone.com/']
match_tokens = build_match_tokens(NAMES, DOMAINS)

print(f"키워드: {KEYWORD}")
print(f"매칭 토큰: {match_tokens}")
print("=" * 70)

# 1. VIDEO 매칭 — fetch_search_page(where='video')
print("\n[1] VIDEO 별도 검색 페이지")
ht_v = fetch_search_page(KEYWORD, where='video')
if not ht_v:
    print("  fetch 실패")
else:
    print(f"  HTML {len(ht_v):,}B")
    cnt_kr = ht_v.count('서울본')
    print(f"  '서울본' 등장 {cnt_kr}회")
    items = extract_video_items_with_dates(ht_v)
    print(f"  추출 동영상 {len(items)}개 — top 10:")
    matched = 0
    for i, (txt, dt) in enumerate(items[:10], 1):
        ntxt = normalize_text(txt)
        is_match = any(t in ntxt for t in match_tokens)
        mark = '★' if is_match else ' '
        if is_match and matched == 0:
            matched = i
        print(f"      {mark}{i}: {txt[:130]} | date={dt}")
    print(f"  → matched_rank = {matched}")

# 2. find_rank_by_web_tab('video') 직접 호출
print("\n[2] find_rank_by_web_tab('video', ...)")
rank, ev = find_rank_by_web_tab('video', KEYWORD, match_tokens)
print(f"  rank={rank}")
print(f"  ev keys: {list(ev.keys())}")
if 'top' in ev:
    for r in ev['top'][:10]:
        print(f"    {r}")

# 3. NAVER 지역검색 API + drt fallback (map)
print("\n[3] MAP — find_rank_by_api_tab('map', ...)")
cid = os.getenv('NAVER_CLIENT_ID')
csec = os.getenv('NAVER_CLIENT_SECRET')
rank, ev = find_rank_by_api_tab('map', KEYWORD, match_tokens, cid, csec)
print(f"  rank={rank}")
print(f"  basis={ev.get('basis')}")
print(f"  primaryBasis={ev.get('primaryBasis')}")
if 'top' in ev:
    print(f"  지역 API top:")
    for r in ev['top'][:10]:
        print(f"    {r.get('rank')}: {r.get('text', '')[:120]}")
if 'drtTop' in ev:
    print(f"  drt fallback top:")
    for r in ev['drtTop']:
        mark = '★' if r['match'] else ' '
        print(f"    {mark}{r['rank']}: {r['name']}")

# 4. 통합검색 페이지의 video 영역 확인 (fallback 단서)
print("\n[4] 통합검색 페이지에서 '서울본' 등장 컨텍스트 (앞뒤 100자, 최대 8개)")
ht_int = fetch_integrated_search_page(KEYWORD)
if ht_int:
    matches = list(re.finditer(r'서울본', ht_int))
    print(f"  통합검색 페이지 {len(ht_int):,}B, '서울본' {len(matches)}회 등장")
    for i, m in enumerate(matches[:8], 1):
        start = max(0, m.start() - 100)
        end = min(len(ht_int), m.end() + 100)
        ctx = ht_int[start:end].replace('\n', ' ').replace('\r', ' ')[:300]
        print(f"      {i}: ...{ctx}...")
else:
    print("  통합검색 fetch 실패")
