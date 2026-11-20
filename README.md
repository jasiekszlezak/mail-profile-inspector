# Mail Profile Inspector

[Instrukcja po polsku](README.pl.md)

A Windows inventory tool for planning workstation and mail client migrations. It lists Thunderbird profiles and their disk usage, Outlook Classic MAPI profiles, and PST/OST files found in known local directories. PST/OST files are listed separately from MAPI profiles. For New Outlook, the scan reports application detection.

The scan reads profile configuration and file metadata. CSV reports include local paths and profile names. The PDF report is a summary without file paths. The CLI and reports currently display Polish labels.

## Run

Use Python 3.10 or newer. Run the scanner on Windows as the current user. The synthetic demo runs on other systems as well.

```sh
python mail_profile_inspector.py --demo
python mail_profile_inspector.py --csv report.csv
```

For PDF export:

```sh
python -m pip install -r requirements-pdf.txt
python mail_profile_inspector.py --demo --csv demo.csv --pdf demo.pdf
```

Drop `--demo` when scanning the current Windows account. The demo uses synthetic entries and can be used to see the report structure before running a local scan.

## Tests

```sh
python -m pip install -r requirements-pdf.txt
python -m unittest discover -s tests -v
```

The scanner was extracted from the JNL Software EML/MSG/PDF project. Developed by [JNL Software](https://www.jnlsoftware.pl/).
