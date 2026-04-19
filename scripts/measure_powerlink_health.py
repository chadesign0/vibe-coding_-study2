# -*- coding: utf-8 -*-
"""
파워링크(순위) 파이프라인 건강도 측정.

- 골든 라벨이 없으면 '정답률'이 아니라 아래 지표만 산출합니다.
  - HTTP 수집 성공률
  - HTML 내 pcPowerLink 블록 탐지율
  - 블록이 있을 때 병원명 매칭(1~10위) 비율

선택: config/golden_powerlink.json 이 있으면 정답 순위와 비교해 일치율을 추가 출력합니다.
형식 예:
{
  "hospitalNames": ["포인트병원"],
  "expected": { "풍산역목통증": 1, "주엽역무릎통증": 0 }
}
"""
from __future__ import annotations

import importlib.util
import json
import sys
import time
from pathlib import Path

from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parents[1]
GOLDEN_PATH = ROOT / "config" / "golden_powerlink.json"


def _load_build():
    path = ROOT / "scripts" / "build_april_month.py"
    spec = importlib.util.spec_from_file_location("build_april_month", path)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader
    spec.loader.exec_module(mod)
    return mod


def main() -> None:
    bam = _load_build()
    cfg_path = ROOT / "config" / (sys.argv[1] if len(sys.argv) > 1 else "april_keywords.json")
    cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
    names = cfg.get("hospitalNames") or []
    keywords = [str(k).strip() for k in (cfg.get("keywords") or []) if str(k).strip()]
    if not keywords:
        raise SystemExit("키워드가 없습니다.")

    golden = {}
    if GOLDEN_PATH.exists():
        try:
            g = json.loads(GOLDEN_PATH.read_text(encoding="utf-8"))
            golden = g.get("expected") or {}
            if g.get("hospitalNames"):
                names = g["hospitalNames"]
        except Exception:
            golden = {}

    http_ok = block_found = cand_nonempty = matched = 0
    http_fail = 0
    details: list[tuple[str, str, int, int]] = []  # kw, status, n_cands, rank

    for kw in keywords:
        ht = bam.fetch_search_page(kw)
        if not ht:
            http_fail += 1
            details.append((kw, "http_fail", 0, 0))
            time.sleep(0.25)
            continue
        http_ok += 1
        soup = BeautifulSoup(ht, "html.parser")
        root = soup.select_one("div[id^='pcPowerLink_']")
        if not root:
            details.append((kw, "no_block", 0, 0))
            time.sleep(0.25)
            continue
        block_found += 1
        cands = bam.extract_candidates_powerlink(ht)
        if cands:
            cand_nonempty += 1
        rank = bam.find_rank_in_candidates(cands, names)
        if rank and rank >= 1:
            matched += 1
        details.append((kw, "ok", len(cands), int(rank or 0)))
        time.sleep(0.25)

    n = len(keywords)
    pct = lambda a, d: round(100.0 * a / d, 1) if d else 0.0

    print("=== 파워링크 측정 결과 ===")
    print(f"키워드 수: {n}")
    print(f"HTTP 수집 성공: {http_ok}/{n} ({pct(http_ok, n)}%)")
    print(f"pcPowerLink 블록 탐지: {block_found}/{n} ({pct(block_found, n)}%)")
    if block_found:
        print(f"  └ 블록 내 후보 텍스트 1개 이상: {cand_nonempty}/{block_found} ({pct(cand_nonempty, block_found)}%)")
        print(f"  └ 병원명 매칭(1~10위): {matched}/{block_found} ({pct(matched, block_found)}%)")
    print()
    print("※ 이 수치는 '정답과의 일치율'이 아니라, 수집·파서가 얼마나 자료를 잡는지에 가깝습니다.")
    print("※ 실제 정확도(%)는 사람이 본 순위와 비교하는 골든셋이 필요합니다.")

    if golden:
        ok = miss = 0
        for kw in keywords:
            if kw not in golden:
                continue
            exp = int(golden[kw])
            act = next((d[3] for d in details if d[0] == kw), -1)
            if act == exp:
                ok += 1
            else:
                miss += 1
        checked = ok + miss
        if checked:
            print()
            print(f"골든셋 비교 ({GOLDEN_PATH.name}): {ok}/{checked} 일치 ({pct(ok, checked)}%)")

    print()
    print("--- 키워드별 요약 ---")
    for kw, st, nc, rk in details:
        print(f"  {kw}: {st} | 후보 {nc}개 | 순위 {rk}")


if __name__ == "__main__":
    main()
