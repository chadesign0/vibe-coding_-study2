import inspect
import json
import tempfile
import threading
import unittest
from pathlib import Path
from unittest import mock


with mock.patch.object(threading.Thread, "start", lambda self: None):
    import server


def month(hospital_name, month_label, keywords):
    rows = [[i + 1, "", keyword] for i, keyword in enumerate(keywords)]
    return {
        "hospitalName": hospital_name,
        "monthLabel": month_label,
        "keywordChannels": {keyword: "all" for keyword in keywords},
        "keywordScopes": {keyword: "all" for keyword in keywords},
        "sheets": [{"key": "regional-pc", "rows": rows}],
    }


class ScopedDeletionTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        (self.root / "data").mkdir()
        (self.root / "config").mkdir()
        self.data_path = self.root / "data" / "scoring-data.json"
        self.data_path.write_text(
            json.dumps(
                {
                    "months": [
                        month("대찬병원", "5월", ["공통키워드", "5월전용"]),
                        month("대찬병원", "6월", ["공통키워드", "6월전용"]),
                        month("두발로병원", "5월", ["공통키워드"]),
                    ]
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        self.data_patch = mock.patch.object(server, "DATA_PATH", self.data_path)
        self.root_patch = mock.patch.object(server, "ROOT", self.root)
        self.data_patch.start()
        self.root_patch.start()

    def tearDown(self):
        self.root_patch.stop()
        self.data_patch.stop()
        self.temp_dir.cleanup()

    def test_scoring_data_deletes_only_selected_hospital_and_month(self):
        changed, removed = server.delete_keyword_from_scoring_data(
            "공통키워드", "5월", "대찬병원"
        )

        self.assertTrue(changed)
        self.assertEqual(removed, 1)
        saved = json.loads(self.data_path.read_text(encoding="utf-8"))
        by_scope = {
            (item["hospitalName"], item["monthLabel"]): [
                row[2] for sheet in item["sheets"] for row in sheet["rows"]
            ]
            for item in saved["months"]
        }
        self.assertNotIn("공통키워드", by_scope[("대찬병원", "5월")])
        self.assertIn("공통키워드", by_scope[("대찬병원", "6월")])
        self.assertIn("공통키워드", by_scope[("두발로병원", "5월")])

    def test_evidence_is_kept_when_another_month_still_uses_keyword(self):
        evidence_path = self.root / "data" / "last-run-evidence.json"
        evidence_path.write_text(
            json.dumps({"evidence": {}, "byHospital": {"대찬병원": {"공통키워드": {"web": {}}}}}),
            encoding="utf-8",
        )
        server.delete_keyword_from_scoring_data("공통키워드", "5월", "대찬병원")

        changed = server.delete_keyword_from_evidence("공통키워드", "대찬병원")

        self.assertFalse(changed)
        saved = json.loads(evidence_path.read_text(encoding="utf-8"))
        self.assertIn("공통키워드", saved["byHospital"]["대찬병원"])

    def test_base_config_deletion_is_scoped(self):
        may_path = self.root / "config" / "may.json"
        june_path = self.root / "config" / "june.json"
        may_path.write_text(
            json.dumps({"hospitalName": "대찬병원", "monthLabel": "5월", "keywords": ["공통키워드"]}),
            encoding="utf-8",
        )
        june_path.write_text(
            json.dumps({"hospitalName": "대찬병원", "monthLabel": "6월", "keywords": ["공통키워드"]}),
            encoding="utf-8",
        )
        missing = self.root / "config" / "missing.json"
        with mock.patch.multiple(
            server,
            CONFIG_PATH=may_path,
            ZEROPAIN_CONFIG_PATH=june_path,
            SAMSUNGBON_CONFIG_PATH=missing,
            JL_CONFIG_PATH=missing,
            SNU_CONFIG_PATH=missing,
        ):
            changed, remaining = server.delete_keyword("공통키워드", "5월", "대찬병원")

        self.assertTrue(changed)
        self.assertEqual(remaining, 0)
        self.assertEqual(json.loads(may_path.read_text(encoding="utf-8"))["keywords"], [])
        self.assertEqual(
            json.loads(june_path.read_text(encoding="utf-8"))["keywords"], ["공통키워드"]
        )


class ScoringSafetyTests(unittest.TestCase):
    def tearDown(self):
        with server.SCORE_TASKS_LOCK:
            server.SCORE_TASKS.clear()
            server.ACTIVE_SCORE_TASK_BY_KEY.clear()

    def test_only_registered_hospitals_are_accepted(self):
        self.assertEqual(server.supported_hospital_name("삼성본정형외과"), "삼성본병원")
        self.assertIsNone(server.supported_hospital_name(None))
        self.assertIsNone(server.supported_hospital_name('$(echo "unsafe")'))

    def test_upload_rejects_unregistered_hospital_before_config_write(self):
        client = server.app.test_client()
        with mock.patch.object(server, "update_keywords") as update:
            response = client.post(
                "/api/upload-keywords",
                data={
                    "keywords_text": "테스트",
                    "hospital_name": '$(echo "unsafe")',
                    "month_label": "5월",
                },
            )

        self.assertEqual(response.status_code, 400)
        update.assert_not_called()

    def test_upload_rejects_same_scope_while_scoring(self):
        with server.SCORE_TASKS_LOCK:
            server.SCORE_TASKS["active-task"] = {
                "taskId": "active-task",
                "status": "running",
                "hospitalName": "대찬병원",
                "monthLabel": "5월",
            }
        client = server.app.test_client()
        with mock.patch.object(server, "update_keywords") as update:
            response = client.post(
                "/api/upload-keywords",
                data={
                    "keywords_text": "테스트",
                    "hospital_name": "대찬병원",
                    "month_label": "5월",
                },
            )

        self.assertEqual(response.status_code, 409)
        update.assert_not_called()

    def test_runtime_config_name_is_content_addressed(self):
        first = server.runtime_config_path_for_config({"hospitalName": "대찬병원", "monthLabel": "5월"})
        same = server.runtime_config_path_for_config({"monthLabel": "5월", "hospitalName": "대찬병원"})
        different = server.runtime_config_path_for_config({"hospitalName": "대찬병원", "monthLabel": "6월"})

        self.assertEqual(first.name, same.name)
        self.assertNotEqual(first.name, different.name)
        self.assertRegex(first.name, r"^runtime_[0-9a-f]{16}\.json$")

    def test_runtime_config_is_pushed_before_workflow_dispatch(self):
        source = inspect.getsource(server.enqueue_actions_rescore_task)
        self.assertLess(
            source.index("_github_push_runtime_config(config_name)"),
            source.index("_trigger_github_actions_workflow("),
        )

    def test_workflow_inputs_are_not_expanded_inside_shell_scripts(self):
        workflow = Path(server.ROOT, ".github", "workflows", "scoring.yml").read_text(encoding="utf-8")
        in_run_block = False
        for line in workflow.splitlines():
            if line.startswith("        run: |"):
                in_run_block = True
                continue
            if in_run_block and line.strip() and not line.startswith("          "):
                in_run_block = False
            if in_run_block:
                self.assertNotIn("${{ inputs.", line)


if __name__ == "__main__":
    unittest.main()
