"""Inventory local mail profiles and data files from the command line."""

from argparse import ArgumentParser
from datetime import datetime
import os
from pathlib import Path

from mail_profile_reports import export_csv_report, export_pdf_report
from mail_profile_scanner import (
    MailProfileEntry,
    MailProfileScanResult,
    MailProfileScanner,
    MailProgramResult,
)


def demo_result() -> MailProfileScanResult:
    demo_home = Path("C:/Users/Demo")
    thunderbird = MailProgramResult(
        "Thunderbird", True, "Przykład", entries=[
            MailProfileEntry(
                "Thunderbird", "Demo", "Profil Thunderbird",
                demo_home / "AppData/Roaming/Thunderbird/Profiles/demo.default",
                4 * 1024 * 1024, "OK", is_default=True,
            ),
        ],
    )
    outlook = MailProgramResult(
        "Outlook Classic", True, "Przykład", entries=[
            MailProfileEntry(
                "Outlook Classic", "Archiwum.pst", "Plik PST",
                demo_home / "Documents/Outlook Files/Archiwum.pst",
                16 * 1024 * 1024, "OK",
            ),
        ],
    )
    return MailProfileScanResult(datetime.now(), [thunderbird, outlook])


def main() -> None:
    parser = ArgumentParser(description=__doc__)
    parser.add_argument("--demo", action="store_true", help="use synthetic data")
    parser.add_argument("--csv", type=Path, metavar="FILE", help="export a CSV report")
    parser.add_argument("--pdf", type=Path, metavar="FILE", help="export a PDF report")
    args = parser.parse_args()

    if os.name != "nt" and not args.demo:
        parser.error("Local scanning requires Windows. Use --demo to try the reports.")
    if args.csv and args.pdf and args.csv.resolve() == args.pdf.resolve():
        parser.error("Choose different paths for the CSV and PDF reports.")

    try:
        result = demo_result() if args.demo else MailProfileScanner().scan()
        for program in result.programs:
            print(f"{program.program}: {program.status} ({program.display_total})")
            for entry in program.entries:
                print(f"  {entry.name} | {entry.item_type} | {entry.display_size} | {entry.path or '—'}")
        if args.csv:
            print(f"CSV: {export_csv_report(result, args.csv)}")
        if args.pdf:
            print(f"PDF: {export_pdf_report(result, args.pdf)}")
    except ImportError as error:
        parser.exit(1, f"{error}. Install PDF dependencies with: python -m pip install -r requirements-pdf.txt\n")
    except OSError as error:
        parser.exit(1, f"{error}\n")


if __name__ == "__main__":
    main()
