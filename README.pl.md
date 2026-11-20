# Mail Profile Inspector

[Wersja angielska](README.md)

Narzędzie do inwentaryzacji profili pocztowych przed migracją stanowiska Windows. Wykrywa profile Thunderbirda i oblicza zajęte przez nie miejsce. Pokazuje konfiguracje MAPI Outlooka Classic oraz osobno pliki PST i OST znalezione w znanych lokalizacjach. Dla Nowego Outlooka raportuje wykrycie aplikacji.

Skanowanie odczytuje konfigurację profili i metadane plików. Raport CSV zawiera lokalne ścieżki i nazwy profili. PDF zawiera podsumowanie bez ścieżek do plików.

## Uruchomienie

Wymagany Python 3.10 lub nowszy. Skanowanie lokalnego konta uruchamia się na Windows. Tryb demonstracyjny działa również na innych systemach.

```sh
python mail_profile_inspector.py --demo
python mail_profile_inspector.py --csv report.csv
```

Aby wygenerować PDF:

```sh
python -m pip install -r requirements-pdf.txt
python mail_profile_inspector.py --demo --csv demo.csv --pdf demo.pdf
```

Uruchom polecenie bez `--demo`, aby użyć ustawień bieżącego konta Windows. Przed przekazaniem CSV sprawdź, czy zawarte w nim ścieżki i nazwy profili mogą trafić do odbiorcy.

## Testy

```sh
python -m pip install -r requirements-pdf.txt
python -m unittest discover -s tests -v
```
