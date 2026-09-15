# -*- coding: utf-8 -*-
"""GitHub Actions 전용: build_month.py 가 임시 파일로 남긴 결과를 최신 데이터 파일에 병합.

push 가 거절되면 rebase 로 JSON 을 섞지 않고, 최신 main / scoring-evidence 를 다시 받은 뒤
이 스크립트로 병합을 새로 수행한다. 다른 병원 채점이 동시에 끝나도 서로의 결과가 유실되지 않는다.

사용법:
  python scripts/merge_results.py scoring  <month.json>
  python scripts/merge_results.py evidence <evidence.json> <target.json>
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_month import ROOT, merge_into_scoring_data, save_evidence  # noqa: E402


def main(argv: list[str]) -> None:
    if len(argv) == 2 and argv[0] == "scoring":
        month = json.loads(Path(argv[1]).read_text(encoding="utf-8"))
        merge_into_scoring_data(month, ROOT / "data" / "scoring-data.json")
    elif len(argv) == 3 and argv[0] == "evidence":
        tmp = json.loads(Path(argv[1]).read_text(encoding="utf-8"))
        save_evidence(tmp.get("evidence") or {}, tmp.get("hospitalName"), Path(argv[2]))
    else:
        raise SystemExit(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
