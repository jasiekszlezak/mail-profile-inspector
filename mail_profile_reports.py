"""Eksport prostych, wyłącznie diagnostycznych raportów skanera profili."""
from __future__ import annotations

import csv
from html import escape
from pathlib import Path
import os
import unicodedata

from mail_profile_scanner import MailProfileScanResult, format_bytes


def _report_rows(result: MailProfileScanResult):
    for program in result.programs:
        if not program.entries:
            yield {
                "Program": program.program,
                "Pozycja": "—",
                "Typ": "—",
                "Zajętość": "—",
                "Status": program.status,
                "Ścieżka": str(program.executable_path or ""),
            }
        for entry in program.entries:
            yield {
                "Program": program.program,
                "Pozycja": entry.name + (" (domyślny)" if entry.is_default else ""),
                "Typ": entry.item_type,
                "Zajętość": format_bytes(entry.size_bytes),
                "Status": entry.status,
                "Ścieżka": str(entry.path or ""),
            }


def _spreadsheet_text(value: str) -> str:
    # Excel may interpret names or statuses beginning with these characters as formulas.
    if value.lstrip().startswith(("=", "+", "-", "@")) or value.startswith(("\t", "\r", "\n")):
        return "'" + value
    return value


def export_csv_report(result: MailProfileScanResult, destination: str | Path) -> Path:
    path = Path(destination)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["Program", "Pozycja", "Typ", "Zajętość", "Status", "Ścieżka"],
            delimiter=";",
        )
        writer.writeheader()
        writer.writerows(
            {key: _spreadsheet_text(value) for key, value in row.items()}
            for row in _report_rows(result)
        )
    return path


def _find_font() -> Path | None:
    system_root = os.environ.get("SystemRoot")
    candidates = [Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")]
    if system_root:
        candidates.insert(0, Path(system_root) / "Fonts" / "segoeui.ttf")
    return next((candidate for candidate in candidates if candidate.is_file()), None)


def export_pdf_report(result: MailProfileScanResult, destination: str | Path) -> Path:
    """Create a diagnostic PDF using the optional ReportLab dependency."""
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

    path = Path(destination)
    path.parent.mkdir(parents=True, exist_ok=True)
    font_path = _find_font()
    font = "Helvetica"
    if font_path:
        font = "MailProfileReport"
        pdfmetrics.registerFont(TTFont(font, str(font_path)))

    styles = {
        "title": ParagraphStyle("Title", fontName=font, fontSize=16, leading=20, spaceAfter=12),
        "heading": ParagraphStyle("Heading", fontName=font, fontSize=12, leading=16, spaceAfter=6),
        "body": ParagraphStyle("Body", fontName=font, fontSize=10, leading=14, alignment=TA_LEFT),
    }

    def paragraph(text: str, style: str = "body") -> Paragraph:
        if not font_path:
            text = text.translate(str.maketrans({"ł": "l", "Ł": "L", "—": "-"}))
            text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
        return Paragraph(escape(text), styles[style])

    story = [
        paragraph("Raport profili pocztowych", "title"),
        paragraph(f"Data skanowania: {result.generated_at:%d.%m.%Y %H:%M}"),
        paragraph(f"Łączna zajętość danych: {result.display_total}"),
        paragraph("Raport nie zawiera treści wiadomości ani listy folderów."),
    ]
    for program in result.programs:
        story.extend([Spacer(1, 12), paragraph(f"{program.program} — {program.status}", "heading")])
        if not program.entries:
            story.append(paragraph("Brak wykrytych profili lub plików danych."))
        for entry in program.entries:
            default = " · domyślny" if entry.is_default else ""
            story.append(paragraph(
                f"{entry.name}{default} | {entry.item_type} | {entry.display_size} | {entry.status}"
            ))
        for note in program.notes:
            story.append(paragraph(f"Uwaga: {note}"))

    SimpleDocTemplate(
        str(path), pagesize=A4, title="Raport profili pocztowych", author="JNL Software",
    ).build(story)
    return path
