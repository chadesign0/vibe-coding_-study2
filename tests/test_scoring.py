import sys
import unittest
from pathlib import Path
from unittest import mock


SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

import build_month


class ScoringDeviceTests(unittest.TestCase):
    def test_pc_and_mobile_sheets_use_separate_rank_results(self):
        cfg = {
            "hospitalName": "테스트병원",
            "rowsBySheetKey": {
                key: [{"region": "서울", "keyword": "허리통증"}]
                for key, _ in build_month.SHEETS_META
            },
        }
        ranks = {
            "pc": {"허리통증": {"powerlink": 1}},
            "mobile": {"허리통증": {"powerlink": 7}},
        }

        payload = build_month.build_month_payload(cfg, ranks, {})
        rows = {sheet["key"]: sheet["rows"][0] for sheet in payload["sheets"]}

        self.assertEqual(rows["regional-pc"][build_month.POWERLINK_COL], 1)
        self.assertEqual(rows["regional-mob"][build_month.POWERLINK_COL], 7)

    def test_local_search_uses_official_result_limit(self):
        response = mock.Mock(status_code=200)
        response.json.return_value = {"items": []}
        with mock.patch.object(build_month.requests, "get", return_value=response) as get:
            result = build_month.api_search("id", "secret", "local", "허리통증")

        self.assertEqual(result, [])
        params = get.call_args.kwargs["params"]
        self.assertEqual(params["display"], 5)
        self.assertEqual(params["sort"], "random")

    def test_reused_rows_remain_scoped_to_their_device_sheet(self):
        record = {
            "sheets": [
                {"key": "regional-pc", "rows": [[1, "서울", "허리통증", None, None, None, None, 1]]},
                {"key": "regional-mob", "rows": [[1, "서울", "허리통증", None, None, None, None, 8]]},
            ]
        }

        reused = build_month._keyword_rows_by_sheet(record)

        self.assertEqual(reused["regional-pc"]["허리통증"][build_month.POWERLINK_COL], 1)
        self.assertEqual(reused["regional-mob"]["허리통증"][build_month.POWERLINK_COL], 8)

    def test_mobile_search_uses_mobile_host_and_separate_cache(self):
        response = mock.Mock(status_code=200, content=b"mobile-result")
        session = mock.Mock()
        session.get.return_value = response
        build_month._INTEGRATED_HTML_CACHE.clear()
        with (
            mock.patch.object(build_month, "_get_web_session", return_value=session),
            mock.patch.object(build_month.time, "sleep"),
        ):
            result = build_month.fetch_integrated_search_page("허리통증", "mobile")

        self.assertEqual(result, "mobile-result")
        self.assertIn("m.search.naver.com", session.get.call_args.args[0])
        self.assertIn("mobile:허리통증", build_month._INTEGRATED_HTML_CACHE)

    def test_mobile_pass_reuses_device_neutral_api_but_rescores_web_channels(self):
        cfg = {"keywords": ["허리통증"], "hospitalNames": ["테스트병원"]}

        def web_rank(_tab, _kw, _tokens, _period, _blog_ids, device="pc"):
            rank = 1 if device == "pc" else 7
            return rank, {"matched_rank": rank}

        def map_rank(_kw, _tokens, *, primary_basis, device="pc"):
            rank = 2 if device == "pc" else 6
            return rank, {"matched_rank": rank, "basis": primary_basis}

        with (
            mock.patch.dict(build_month.os.environ, {"PLAYWRIGHT_VERIFY_ENABLED": "0"}),
            mock.patch.object(build_month, "find_rank_by_web_tab", side_effect=web_rank),
            mock.patch.object(build_month, "_try_map_drt_fallback", side_effect=map_rank),
            mock.patch.object(build_month, "find_rank_by_api_tab", return_value=(3, {"matched_rank": 3})),
        ):
            pc_ranks, pc_evidence = build_month.fetch_keyword_ranks(cfg, "id", "secret", device="pc")
            mobile_ranks, _ = build_month.fetch_keyword_ranks(
                cfg,
                "id",
                "secret",
                device="mobile",
                shared_api=(pc_ranks, pc_evidence),
            )

        self.assertEqual(pc_ranks["허리통증"]["web"], 1)
        self.assertEqual(mobile_ranks["허리통증"]["web"], 7)
        self.assertEqual(pc_ranks["허리통증"]["map"], 2)
        self.assertEqual(mobile_ranks["허리통증"]["map"], 6)
        self.assertEqual(pc_ranks["허리통증"]["blog"], mobile_ranks["허리통증"]["blog"])


if __name__ == "__main__":
    unittest.main()
