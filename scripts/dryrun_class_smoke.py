# -*- coding: utf-8 -*-
"""drt fallback + Playwright verify 통합 smoke test.

클래스병원의 keyword 5개로 fetch_keyword_ranks 직접 호출.
- 중앙역정형외과: drt fallback 또는 Playwright로 map 점수 회복 기대
- 안산정형외과: 1차 NAVER 지역 API에서 매칭 가능성
- 본오동정형외과 / 성포동정형외과: 모든 채널 0점 → Playwright 발동 후보
- 안산알츠하이머: 일부 채널만 점수 가능
"""
import io
import os
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, str(Path(__file__).parent))

from dotenv import load_dotenv
load_dotenv()

# parallel을 1로 — throttle 회피
os.environ['SCORING_PARALLEL_WORKERS'] = '1'
os.environ['PLAYWRIGHT_VERIFY_ENABLED'] = '1'

from build_month import fetch_keyword_ranks  # noqa: E402

cid = os.getenv('NAVER_CLIENT_ID')
csec = os.getenv('NAVER_CLIENT_SECRET')
assert cid and csec, "NAVER_CLIENT_ID/SECRET 필요"

cfg = {
    "monthLabel": "5월",
    "hospitalName": "클래스병원",
    "hospitalNames": ["클래스병원"],
    "hospitalDomains": ["https://class2023.co.kr/"],
    "hospitalBlogBases": [
        "https://blog.naver.com/classlim2",
        "https://blog.naver.com/class231004",
    ],
    "keywords": [
        "중앙역정형외과",         # drt fallback 회복 기대
        "안산병원",                # 전 채널 0 — Playwright 발동 후보
        "안산치매",                # 전 채널 0 — Playwright 발동 후보
        "오십견치료방법",          # 전 채널 0 — 광역 일반 키워드
        "본오동정형외과",          # 인접동 (참고)
    ],
    "keywordChannels": {},
    "keywordScopes": {},
    "manualRanks": {},
    "manualRanksByTab": {},
}

print("=" * 70)
print("dry-run smoke test 시작 — 5개 키워드")
print("=" * 70)

out, ev = fetch_keyword_ranks(cfg, cid, csec, report_progress=False)

print("\n" + "=" * 70)
print("최종 결과")
print("=" * 70)
channels = ["powerlink", "bizsite", "map", "cafe", "blog", "news", "video", "web"]
print(f"{'keyword':18s} | " + " | ".join(f"{c[:5]:>5s}" for c in channels))
print("-" * 90)
for kw, ranks in out.items():
    cells = [f"{ranks.get(c, 0) or 0:>5d}" for c in channels]
    print(f"{kw:18s} | " + " | ".join(cells))

print("\n=== drt / playwright 흔적 (evidence) ===")
for kw, ev_kw in ev.items():
    for ch, ev_ch in ev_kw.items():
        if not isinstance(ev_ch, dict):
            continue
        if ev_ch.get("basis") == "integrated_search_drt_fallback":
            print(f"  [drt-hit] {kw}/{ch}: rank={ev_ch.get('matched_rank')} primaryBasis={ev_ch.get('primaryBasis')}")
        elif ev_ch.get("drtFallback"):
            fb = ev_ch["drtFallback"]
            print(f"  [drt-miss] {kw}/{ch}: drtCount={fb.get('drtCount')} primaryBasis={fb.get('primaryBasis')}")
        if ev_ch.get("playwrightVerify"):
            pw = ev_ch["playwrightVerify"]
            print(f"  [pw-hit] {kw}/{ch}: rank={pw.get('rank')} via={pw.get('via')}")
