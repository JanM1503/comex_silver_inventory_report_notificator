import json
import tempfile
import unittest
from datetime import date, datetime, timezone
from pathlib import Path
from unittest.mock import Mock, patch

import silver_alert as app


class ReliabilityTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        for name, value in {
            "STATE_DIR": self.root,
            "LAST_REPORT_FILE": self.root / "last_report.txt",
            "OBS_FILE": self.root / "observations.json",
            "REPORT_DIR": self.root / "reports",
            "TEST_RECEIPT": self.root / "last_test_push.json",
        }.items():
            p = patch.object(app, name, value)
            p.start()
            self.addCleanup(p.stop)
        self.now = datetime(2026, 11, 27, 12, 0, tzinfo=timezone.utc)
        self.rep = app.Report(date(2026, 11, 25), {
            "Total": (100.0, 90.0),
            "Eligible": (60.0, 55.0),
            "Registered": (40.0, 35.0),
        }, "abc", self.now)
        for name, value in {"_now": lambda: self.now, "_sleep": Mock()}.items():
            p = patch.object(app, name, value)
            p.start()
            self.addCleanup(p.stop)

    def entry(self, d="2026-11-25"):
        return json.loads((app.REPORT_DIR / (d + ".json")).read_text())

    def test_every_start_probes_and_duplicate_is_not_sent(self):
        with patch.object(app, "fetch_report", return_value=self.rep) as fetch, patch.object(app, "send_ntfy", return_value="msg-1") as send:
            self.assertEqual(app.run_session("private-topic"), 0)
            self.assertEqual(app.run_session("private-topic"), 0)
            self.assertEqual(fetch.call_count, 2)
            send.assert_called_once()
        e = self.entry()
        self.assertEqual(e["inventory"]["Total"]["change"], -10.0)
        self.assertEqual(e["delivery"]["ntfy_id"], "msg-1")
        self.assertEqual(e["delivery"]["status"], "accepted")
        self.assertEqual(app.notified_date(), self.rep.report_date)
        self.assertNotIn("private-topic", json.dumps(e))

    def test_failed_send_keeps_report_and_retries_even_if_cme_unavailable(self):
        with patch.object(app, "fetch_report", return_value=self.rep), patch.object(app, "send_ntfy", side_effect=OSError("down")):
            self.assertEqual(app.run_session("topic"), 1)
        self.assertEqual(self.entry()["delivery"]["status"], "pending")
        self.assertIsNone(app.notified_date())
        with patch.object(app, "fetch_report", side_effect=OSError("403")) as fetch, patch.object(app, "send_ntfy", return_value="retry-id"):
            self.assertEqual(app.run_session("topic"), 1)
            self.assertEqual(fetch.call_count, 4)
        self.assertEqual(self.entry()["delivery"]["ntfy_id"], "retry-id")
        self.assertEqual(self.entry()["delivery"]["attempts"], 2)

    def test_legacy_report_archived_without_duplicate(self):
        app.mark_notified(self.rep.report_date)
        with patch.object(app, "fetch_report", return_value=self.rep), patch.object(app, "send_ntfy") as send:
            self.assertEqual(app.run_session("topic"), 0)
            send.assert_not_called()
        self.assertEqual(self.entry()["delivery"]["status"], "legacy_cutoff")

    def test_pending_reports_sent_in_order(self):
        older = app.Report(date(2026, 11, 24), self.rep.data, "old", self.now)
        app.archive_report(older, self.now)
        with patch.object(app, "fetch_report", return_value=self.rep), patch.object(app, "send_ntfy", side_effect=["old-id", "new-id"]) as send:
            self.assertEqual(app.run_session("topic"), 0)
            self.assertIn("24.11.2026", send.call_args_list[0].args[1])
            self.assertIn("25.11.2026", send.call_args_list[1].args[1])
        self.assertEqual(self.entry("2026-11-24")["delivery"]["ntfy_id"], "old-id")
        self.assertEqual(self.entry()["delivery"]["ntfy_id"], "new-id")

    def test_revised_report_keeps_original_delivery(self):
        with patch.object(app, "fetch_report", return_value=self.rep), patch.object(app, "send_ntfy", return_value="original") as send:
            app.run_session("topic")
            self.rep.digest = "revised"
            app.run_session("topic")
            app.run_session("topic")
            send.assert_called_once()
        e = self.entry()
        self.assertEqual(e["delivery"]["ntfy_id"], "original")
        self.assertEqual(len(e["revisions"]), 1)
        self.assertEqual(e["revisions"][0]["sha256"], "revised")

    def test_test_push_is_separate_and_does_not_advance_marker(self):
        with patch.object(app, "fetch_report", return_value=self.rep), patch.object(app, "send_ntfy", return_value="test-id") as send:
            self.assertEqual(app.test_push("topic"), 0)
            self.assertTrue(send.call_args.args[1].startswith("TEST - "))
        self.assertFalse(app.LAST_REPORT_FILE.exists())
        self.assertFalse(app.REPORT_DIR.exists())
        self.assertEqual(json.loads(app.TEST_RECEIPT.read_text())["delivery"]["ntfy_id"], "test-id")

    def test_ntfy_response_validation(self):
        response = Mock()
        response.__enter__ = Mock(return_value=response)
        response.__exit__ = Mock(return_value=False)
        with patch.object(app.urllib.request, "urlopen", return_value=response):
            response.read.return_value = b'{"id":"abcdef","topic":"do-not-store","event":"message"}'
            self.assertEqual(app.send_ntfy("topic", "title", "body"), "abcdef")
            for payload in (b"not json", b"{}", b"[]", b'{"id":""}'):
                response.read.return_value = payload
                with self.assertRaises(OSError):
                    app.send_ntfy("topic", "title", "body")

    def test_transient_download_failure_then_success(self):
        with patch.object(app, "fetch_report", side_effect=[OSError("403"), self.rep]), patch.object(app, "send_ntfy", return_value="ok"):
            self.assertEqual(app.run_session("topic"), 0)
        app._sleep.assert_called_once_with(30)

    def test_accepted_receipt_repairs_marker_without_resend(self):
        with patch.object(app, "fetch_report", return_value=self.rep), patch.object(app, "send_ntfy", return_value="ok"):
            self.assertEqual(app.run_session("topic"), 0)
        app.LAST_REPORT_FILE.unlink()
        with patch.object(app, "fetch_report", return_value=self.rep), patch.object(app, "send_ntfy") as send:
            self.assertEqual(app.run_session("topic"), 0)
            send.assert_not_called()
        self.assertEqual(app.notified_date(), self.rep.report_date)


if __name__ == "__main__":
    unittest.main()
