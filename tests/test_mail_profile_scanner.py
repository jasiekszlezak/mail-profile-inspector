from __future__ import annotations

from configparser import ConfigParser
from pathlib import Path
import tempfile
import unittest

from mail_profile_scanner import MailProfileScanner, ScannerPaths, format_bytes


class MailProfileScannerTests(unittest.TestCase):
    def _paths(self, root: Path) -> ScannerPaths:
        return ScannerPaths(
            appdata=root / "AppData" / "Roaming",
            localappdata=root / "AppData" / "Local",
            userprofile=root,
            program_files=root / "Program Files",
            program_files_x86=root / "Program Files (x86)",
        )

    def test_thunderbird_profiles_ini_detects_default_profile_and_size(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            paths = self._paths(root)
            thunderbird = paths.appdata / "Thunderbird"
            profile = thunderbird / "Profiles" / "abc.default-release"
            profile.mkdir(parents=True)
            (profile / "mailbox").write_bytes(b"x" * 1536)
            config = ConfigParser()
            config["Profile0"] = {"Name": "Poczta firmowa", "IsRelative": "1", "Path": "Profiles/abc.default-release", "Default": "1"}
            thunderbird.mkdir(parents=True, exist_ok=True)
            with (thunderbird / "profiles.ini").open("w", encoding="utf-8") as handle:
                config.write(handle)

            result = MailProfileScanner(paths).scan_thunderbird()

            self.assertEqual(len(result.entries), 1)
            entry = result.entries[0]
            self.assertEqual(entry.name, "Poczta firmowa")
            self.assertTrue(entry.is_default)
            self.assertEqual(entry.size_bytes, 1536)
            self.assertEqual(entry.status, "OK")

    def test_thunderbird_falls_back_to_profile_folders_without_ini(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            paths = self._paths(root)
            profile = paths.appdata / "Thunderbird" / "Profiles" / "orphan.default"
            profile.mkdir(parents=True)
            (profile / "data").write_bytes(b"x" * 10)

            result = MailProfileScanner(paths).scan_thunderbird()

            self.assertEqual(result.entries[0].name, "orphan.default")
            self.assertIn("profiles.ini", result.notes[0])

    def test_outlook_data_files_are_listed_without_false_profile_assignment(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            paths = self._paths(root)
            local_data = paths.localappdata / "Microsoft" / "Outlook"
            docs_data = paths.userprofile / "Documents" / "Outlook Files"
            local_data.mkdir(parents=True)
            docs_data.mkdir(parents=True)
            (local_data / "konto.ost").write_bytes(b"x" * 2048)
            (docs_data / "archiwum.pst").write_bytes(b"x" * 4096)

            scanner = MailProfileScanner(paths)
            scanner._outlook_profile_names = lambda: ["Praca"]  # type: ignore[method-assign]
            result = scanner.scan_outlook_classic()

            profile_entries = [entry for entry in result.entries if entry.item_type == "Profil MAPI"]
            data_entries = [entry for entry in result.entries if "plik danych Outlooka" in entry.item_type]
            self.assertEqual(profile_entries[0].name, "Praca")
            self.assertEqual({entry.name for entry in data_entries}, {"konto.ost", "archiwum.pst"})
            self.assertEqual(sum(entry.size_bytes or 0 for entry in data_entries), 6144)

    def test_format_bytes_is_polish_readable(self) -> None:
        self.assertEqual(format_bytes(None), "—")
        self.assertEqual(format_bytes(1024), "1,0 KB")
        self.assertEqual(format_bytes(1536), "1,5 KB")


if __name__ == "__main__":
    unittest.main(verbosity=2)
