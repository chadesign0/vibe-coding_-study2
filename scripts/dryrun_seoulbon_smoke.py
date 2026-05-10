# -*- coding: utf-8 -*-
"""서울본정형외과 video fallback + 통합검색 캐시 동작 검증.

사용자 보고: '처인구재활치료' video page 1위에 서울본정형외과 영상이 노출되는데
5/8 채점에서 동영상=0. video page fetch 실패 또는 throttle 추정.

이번 PR의 video 통합검색 fallback이 작동하는지 dry-run.
"""
import io
import os
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, str(Path(__file__).parent))

from dotenv import load_dotenv
load_dotenv()

os.environ['SCORING_PARALLEL_WORKERS'] = '1'
os.environ['PLAYWRIGHT_VERIFY_ENABLED'] = '1'

from build_month import fetch_keyword_ranks  # noqa: E402

cid = os.getenv('NAVER_CLIENT_ID')
csec = os.getenv('NAVER_CLIENT_SECRET')
assert cid and csec, "NAVER_CLIENT_ID/SECRET 필요"

cfg = {
    "monthLabel": "5월",
    "hospitalName": "서울본정형외과",
    "hospitalNames": ["서울본정형외과"],
    "hospitalDomains": ["https://yonginseoulbone.com/"],
    "hospitalBlogBases": ["https://blog.naver.com/yonginseoulbone"],
    "keywords": [
        "처인구재활치료",   # 사용자 지적 — 동영상 1위인데 5/8엔 0점
        "용인정형외과",     # 광역
        "처인구정형외과",   # 광역
    ],
    "keywordChannels": {},
    "keywordScopes": {},
    "manualRanks": {},
    "manualRanksByTab": {},
}

print("=" * 70)
print("dry-run: 서울본정형외과 — video fallback 검증")
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

print("\n=== fallback / playwright 흔적 ===")
for kw, ev_kw in ev.items():
    for ch, ev_ch in ev_kw.items():
        if not isinstance(ev_ch, dict):
            continue
        basis = ev_ch.get("basis") or ""
        if "fallback" in basis or "drt" in basis:
            print(f"  [{basis}] {kw}/{ch}: rank={ev_ch.get('matched_rank')} primaryBasis={ev_ch.get('primaryBasis')}")
        if ev_ch.get("playwrightVerify"):
            pw = ev_ch["playwrightVerify"]
            print(f"  [pw-hit] {kw}/{ch}: rank={pw.get('rank')} via={pw.get('via')}")
        if ev_ch.get("videoPageFetchFailed"):
            print(f"  [video-fetch-fail] {kw}/{ch}")
