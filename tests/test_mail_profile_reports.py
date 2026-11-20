from __future__ import annotations

from datetime import datetime
from pathlib import Path
import tempfile
import unittest

from mail_profile_reports import export_csv_report, export_pdf_report
from mail_profile_scanner import MailProfileEntry, MailProfileScanResult, MailProgramResult


class MailProfileReportTests(unittest.TestCase):
    def _result(self) -> MailProfileScanResult:
        thunderbird = MailProgramResult(
            "Thunderbird",
            True,
            "Zainstalowany",
            entries=[
                MailProfileEntry("Thunderbird", "Poczta firmowa", "Profil Thunderbird", Path("C:/profile"), 2048, "OK", is_default=True)
            ],
        )
        return MailProfileScanResult(datetime(2026, 7, 2, 12, 0), [thunderbird])

    def test_csv_contains_header_and_profile_name(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            path = export_csv_report(self._result(), Path(temp) / "report.csv")
            data = path.read_text(encoding="utf-8-sig")
            self.assertIn("Program;Pozycja", data)
            self.assertIn("Poczta firmowa (domyślny)", data)

    def test_pdf_report_is_created(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            path = export_pdf_report(self._result(), Path(temp) / "report.pdf")
            self.assertTrue(path.is_file())
            self.assertGreater(path.stat().st_size, 500)


if __name__ == "__main__":
    unittest.main(verbosity=2)
