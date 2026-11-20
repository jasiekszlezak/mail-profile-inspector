"""Read-only scanner lokalnych profili pocztowych Thunderbird i Outlook.

Moduł nie otwiera ani nie analizuje treści wiadomości. Odczytuje wyłącznie
konfigurację profili, nazwy plików danych oraz ich rozmiary na dysku.
"""
from __future__ import annotations

from configparser import ConfigParser, Error as ConfigParserError
from dataclasses import dataclass, field
from datetime import datetime
import os
from pathlib import Path
import re
from typing import Callable, Iterable


CANCELLED = "__SCAN_CANCELLED__"


@dataclass(frozen=True)
class MailProfileEntry:
    """Jedna widoczna pozycja raportu: profil, plik danych lub aplikacja."""

    program: str
    name: str
    item_type: str
    path: Path | None
    size_bytes: int | None
    status: str
    is_default: bool = False
    notes: str = ""

    @property
    def display_size(self) -> str:
        return format_bytes(self.size_bytes)


@dataclass
class MailProgramResult:
    """Wynik skanowania jednego programu pocztowego."""

    program: str
    installed: bool
    status: str
    executable_path: Path | None = None
    entries: list[MailProfileEntry] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    @property
    def total_bytes(self) -> int:
        return sum(entry.size_bytes or 0 for entry in self.entries)

    @property
    def display_total(self) -> str:
        return format_bytes(self.total_bytes) if self.total_bytes else "—"


@dataclass
class MailProfileScanResult:
    """Kompletny wynik jednego skanowania lokalnego użytkownika."""

    generated_at: datetime
    programs: list[MailProgramResult]
    cancelled: bool = False

    @property
    def total_bytes(self) -> int:
        return sum(program.total_bytes for program in self.programs)

    @property
    def display_total(self) -> str:
        return format_bytes(self.total_bytes) if self.total_bytes else "—"

    def program(self, name: str) -> MailProgramResult | None:
        return next((program for program in self.programs if program.program == name), None)


@dataclass(frozen=True)
class ScannerPaths:
    """Ścieżki bazowe użytkownika — wydzielone dla testów i działania na Windows."""

    appdata: Path
    localappdata: Path
    userprofile: Path
    program_files: Path
    program_files_x86: Path

    @classmethod
    def current_user(cls) -> "ScannerPaths":
        userprofile = Path(os.environ.get("USERPROFILE") or Path.home())
        appdata = Path(os.environ.get("APPDATA") or (userprofile / "AppData" / "Roaming"))
        localappdata = Path(os.environ.get("LOCALAPPDATA") or (userprofile / "AppData" / "Local"))
        program_files = Path(os.environ.get("ProgramFiles") or r"C:\\Program Files")
        program_files_x86 = Path(os.environ.get("ProgramFiles(x86)") or r"C:\\Program Files (x86)")
        return cls(appdata, localappdata, userprofile, program_files, program_files_x86)


ProgressCallback = Callable[[str], None]
CancelledCallback = Callable[[], bool]


def format_bytes(value: int | None) -> str:
    """Format rozmiaru bez nadmiernej precyzji dla widoku administracyjnego."""
    if value is None:
        return "—"
    value = max(0, int(value))
    units = ("B", "KB", "MB", "GB", "TB")
    amount = float(value)
    for unit in units:
        if amount < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(amount)} B"
            return f"{amount:.1f} {unit}".replace(".", ",")
        amount /= 1024
    return "—"


def _safe_resolve(path: Path) -> Path:
    try:
        return path.resolve()
    except OSError:
        return path.absolute()


def directory_size(path: Path, *, cancelled: CancelledCallback | None = None) -> tuple[int, int]:
    """Zwraca sumę bajtów i liczbę niedostępnych pozycji.

    Nie podąża za dowiązaniami katalogów, aby uniknąć pętli i przypadkowego
    wyjścia poza profil. Nazwy i treść plików nie są odczytywane.
    """
    total = 0
    denied = 0
    stack = [path]
    while stack:
        if cancelled and cancelled():
            raise RuntimeError(CANCELLED)
        current = stack.pop()
        try:
            with os.scandir(current) as entries:
                for entry in entries:
                    if cancelled and cancelled():
                        raise RuntimeError(CANCELLED)
                    try:
                        if entry.is_dir(follow_symlinks=False):
                            stack.append(Path(entry.path))
                        elif entry.is_file(follow_symlinks=False):
                            total += entry.stat(follow_symlinks=False).st_size
                    except OSError:
                        denied += 1
        except OSError:
            denied += 1
    return total, denied


def file_size(path: Path) -> int | None:
    try:
        return path.stat().st_size
    except OSError:
        return None


def _registry_value(root, subkey: str, value_name: str = "") -> str | None:
    try:
        import winreg  # type: ignore

        with winreg.OpenKey(root, subkey) as key:
            value, _ = winreg.QueryValueEx(key, value_name)
            return str(value) if value else None
    except (ImportError, OSError):
        return None


def _registry_subkeys(root, subkey: str) -> list[str]:
    try:
        import winreg  # type: ignore

        names: list[str] = []
        with winreg.OpenKey(root, subkey) as key:
            count = winreg.QueryInfoKey(key)[0]
            for index in range(count):
                names.append(winreg.EnumKey(key, index))
        return names
    except (ImportError, OSError):
        return []


def _app_path_from_registry(executable_name: str) -> Path | None:
    """Odczytuje standardowy klucz App Paths w obu widokach rejestru."""
    try:
        import winreg  # type: ignore
    except ImportError:
        return None

    key_name = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{executable_name}"
    flags = [0]
    for flag_name in ("KEY_WOW64_64KEY", "KEY_WOW64_32KEY"):
        flag = getattr(winreg, flag_name, 0)
        if flag not in flags:
            flags.append(flag)
    for flag in flags:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_name, 0, winreg.KEY_READ | flag) as key:
                value, _ = winreg.QueryValueEx(key, "")
                candidate = Path(str(value).strip('"'))
                if candidate.is_file():
                    return candidate
        except OSError:
            continue
    return None


class MailProfileScanner:
    """Skaner lokalny Thunderbird, Outlook Classic i Nowego Outlooka."""

    THUNDERBIRD = "Thunderbird"
    OUTLOOK_CLASSIC = "Outlook Classic"
    OUTLOOK_NEW = "Nowy Outlook"

    def __init__(
        self,
        paths: ScannerPaths | None = None,
        *,
        cancelled: CancelledCallback | None = None,
        progress: ProgressCallback | None = None,
    ) -> None:
        self.paths = paths or ScannerPaths.current_user()
        self._cancelled = cancelled or (lambda: False)
        self._progress = progress or (lambda _message: None)

    def scan(self) -> MailProfileScanResult:
        programs: list[MailProgramResult] = []
        try:
            programs.append(self.scan_thunderbird())
            programs.append(self.scan_outlook_classic())
            programs.append(self.scan_new_outlook())
            return MailProfileScanResult(datetime.now(), programs)
        except RuntimeError as exc:
            if str(exc) == CANCELLED:
                return MailProfileScanResult(datetime.now(), programs, cancelled=True)
            raise

    def _check_cancelled(self) -> None:
        if self._cancelled():
            raise RuntimeError(CANCELLED)

    def _emit(self, message: str) -> None:
        self._check_cancelled()
        self._progress(message)

    def _find_thunderbird_exe(self) -> Path | None:
        candidates = [
            _app_path_from_registry("thunderbird.exe"),
            self.paths.program_files / "Mozilla Thunderbird" / "thunderbird.exe",
            self.paths.program_files_x86 / "Mozilla Thunderbird" / "thunderbird.exe",
        ]
        return next((path for path in candidates if path and path.is_file()), None)

    def scan_thunderbird(self) -> MailProgramResult:
        self._emit("Sprawdzanie profili Mozilla Thunderbird…")
        root = self.paths.appdata / "Thunderbird"
        profiles_ini = root / "profiles.ini"
        profiles_root = root / "Profiles"
        executable = self._find_thunderbird_exe()
        entries: list[MailProfileEntry] = []
        notes: list[str] = []

        profiles: list[tuple[str, Path, bool]] = []
        if profiles_ini.is_file():
            parser = ConfigParser(interpolation=None)
            try:
                parser.read(profiles_ini, encoding="utf-8-sig")
                for section in parser.sections():
                    if not section.lower().startswith("profile"):
                        continue
                    name = parser.get(section, "Name", fallback=section).strip() or section
                    raw_path = parser.get(section, "Path", fallback="").strip()
                    is_relative = parser.get(section, "IsRelative", fallback="1").strip() != "0"
                    if not raw_path:
                        notes.append(f"{name}: brak ścieżki w profiles.ini")
                        continue
                    candidate = (root / raw_path) if is_relative else Path(raw_path)
                    is_default = parser.get(section, "Default", fallback="0").strip() == "1"
                    profiles.append((name, candidate, is_default))
            except (OSError, ConfigParserError, UnicodeError) as exc:
                notes.append(f"Nie można odczytać profiles.ini: {exc}")
        elif profiles_root.is_dir():
            notes.append("Nie znaleziono profiles.ini — pokazano wykryte foldery profili.")
            try:
                for child in sorted(profiles_root.iterdir(), key=lambda item: item.name.lower()):
                    if child.is_dir():
                        profiles.append((child.name, child, False))
            except OSError as exc:
                notes.append(f"Nie można odczytać katalogu Profiles: {exc}")

        seen_paths: set[Path] = set()
        for name, profile_path, is_default in profiles:
            self._emit(f"Thunderbird: liczenie profilu {name}…")
            resolved = _safe_resolve(profile_path)
            if resolved in seen_paths:
                continue
            seen_paths.add(resolved)
            if not profile_path.is_dir():
                entries.append(
                    MailProfileEntry(
                        self.THUNDERBIRD,
                        name,
                        "Profil Thunderbird",
                        profile_path,
                        None,
                        "Folder profilu niedostępny",
                        is_default=is_default,
                    )
                )
                continue
            size, denied = directory_size(profile_path, cancelled=self._cancelled)
            status = "OK" if not denied else f"OK · brak dostępu do {denied} pozycji"
            entries.append(
                MailProfileEntry(
                    self.THUNDERBIRD,
                    name,
                    "Profil Thunderbird",
                    profile_path,
                    size,
                    status,
                    is_default=is_default,
                )
            )

        installed = bool(executable and executable.is_file())
        if installed:
            status = "Zainstalowany"
        elif entries:
            status = "Dane profili wykryte · program niewykryty"
        else:
            status = "Nie wykryto"
        return MailProgramResult(self.THUNDERBIRD, installed, status, executable, entries, notes)

    def _find_outlook_exe(self) -> Path | None:
        candidates = [
            _app_path_from_registry("OUTLOOK.EXE"),
            self.paths.program_files / "Microsoft Office" / "root" / "Office16" / "OUTLOOK.EXE",
            self.paths.program_files_x86 / "Microsoft Office" / "root" / "Office16" / "OUTLOOK.EXE",
            self.paths.program_files / "Microsoft Office" / "Office16" / "OUTLOOK.EXE",
            self.paths.program_files_x86 / "Microsoft Office" / "Office16" / "OUTLOOK.EXE",
        ]
        return next((path for path in candidates if path and path.is_file()), None)

    def _outlook_profile_names(self) -> list[str]:
        """Odczytuje nazwy konfiguracji MAPI — bez ich modyfikowania."""
        try:
            import winreg  # type: ignore
        except ImportError:
            return []

        base = r"Software\Microsoft\Office"
        versions = _registry_subkeys(winreg.HKEY_CURRENT_USER, base)
        names: set[str] = set()
        for version in versions:
            if not re.fullmatch(r"\d+\.\d+", version):
                continue
            path = base + rf"\{version}\Outlook\Profiles"
            names.update(_registry_subkeys(winreg.HKEY_CURRENT_USER, path))
        # Dodatkowa ścieżka starszego podsystemu MAPI.
        names.update(
            _registry_subkeys(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles",
            )
        )
        return sorted((name for name in names if name.strip()), key=str.casefold)

    def _force_ost_paths(self) -> list[Path]:
        try:
            import winreg  # type: ignore
        except ImportError:
            return []
        base = r"Software\Microsoft\Office"
        paths: list[Path] = []
        for version in _registry_subkeys(winreg.HKEY_CURRENT_USER, base):
            if not re.fullmatch(r"\d+\.\d+", version):
                continue
            value = _registry_value(winreg.HKEY_CURRENT_USER, base + rf"\{version}\Outlook", "ForceOSTPath")
            if value:
                expanded = os.path.expandvars(value)
                paths.append(Path(expanded))
        return paths

    def _outlook_data_directories(self) -> list[Path]:
        directories = [
            self.paths.localappdata / "Microsoft" / "Outlook",
            self.paths.userprofile / "Documents" / "Outlook Files",
            self.paths.appdata / "Microsoft" / "Outlook",
        ]
        directories.extend(self._force_ost_paths())
        unique: list[Path] = []
        seen: set[Path] = set()
        for directory in directories:
            resolved = _safe_resolve(directory)
            if resolved not in seen:
                unique.append(directory)
                seen.add(resolved)
        return unique

    def _find_outlook_data_files(self) -> list[Path]:
        found: list[Path] = []
        seen: set[Path] = set()
        for directory in self._outlook_data_directories():
            self._emit(f"Outlook: sprawdzanie danych w {directory.name or directory.drive}…")
            if not directory.is_dir():
                continue
            stack = [directory]
            while stack:
                self._check_cancelled()
                current = stack.pop()
                try:
                    with os.scandir(current) as items:
                        for item in items:
                            self._check_cancelled()
                            try:
                                if item.is_dir(follow_symlinks=False):
                                    stack.append(Path(item.path))
                                elif item.is_file(follow_symlinks=False) and Path(item.name).suffix.lower() in {".pst", ".ost"}:
                                    path = Path(item.path)
                                    resolved = _safe_resolve(path)
                                    if resolved not in seen:
                                        found.append(path)
                                        seen.add(resolved)
                            except OSError:
                                continue
                except OSError:
                    continue
        return sorted(found, key=lambda item: item.name.casefold())

    def scan_outlook_classic(self) -> MailProgramResult:
        self._emit("Sprawdzanie Outlook Classic i plików PST / OST…")
        executable = self._find_outlook_exe()
        names = self._outlook_profile_names()
        data_files = self._find_outlook_data_files()
        entries: list[MailProfileEntry] = []

        for name in names:
            entries.append(
                MailProfileEntry(
                    self.OUTLOOK_CLASSIC,
                    name,
                    "Profil MAPI",
                    None,
                    None,
                    "Konfiguracja profilu",
                    is_default=False,
                    notes="Rozmiar danych Outlooka pokazano osobno dla plików PST/OST.",
                )
            )
        for path in data_files:
            self._emit(f"Outlook: odczyt rozmiaru {path.name}…")
            size = file_size(path)
            entries.append(
                MailProfileEntry(
                    self.OUTLOOK_CLASSIC,
                    path.name,
                    path.suffix.upper().lstrip(".") + " — plik danych Outlooka",
                    path,
                    size,
                    "OK" if size is not None else "Brak dostępu do pliku",
                )
            )

        installed = bool(executable and executable.is_file())
        if installed:
            status = "Zainstalowany"
        elif names or data_files:
            status = "Profile lub dane wykryte · program niewykryty"
        else:
            status = "Nie wykryto"
        notes = [
            "Nazwy profili Outlooka są konfiguracją MAPI. System nie gwarantuje bezpiecznego przypisania każdego pliku PST/OST do konkretnego profilu, dlatego pliki danych są pokazane osobno.",
        ]
        return MailProgramResult(self.OUTLOOK_CLASSIC, installed, status, executable, entries, notes)

    def _find_new_outlook_package(self) -> Path | None:
        packages = self.paths.localappdata / "Packages"
        if not packages.is_dir():
            return None
        try:
            matches = sorted(
                (item for item in packages.iterdir() if item.is_dir() and item.name.lower().startswith("microsoft.outlookforwindows_")),
                key=lambda item: item.name.casefold(),
            )
            return matches[0] if matches else None
        except OSError:
            return None

    def scan_new_outlook(self) -> MailProgramResult:
        self._emit("Sprawdzanie Nowego Outlooka…")
        package = self._find_new_outlook_package()
        if package:
            entry = MailProfileEntry(
                self.OUTLOOK_NEW,
                "Nowy Outlook dla Windows",
                "Aplikacja Microsoft Store",
                package,
                None,
                "Wykryto aplikację",
                notes="Nowy Outlook nie korzysta z klasycznych profili MAPI ani standardowych plików OST/PST jak Outlook Classic.",
            )
            return MailProgramResult(
                self.OUTLOOK_NEW,
                True,
                "Zainstalowany",
                None,
                [entry],
                ["Dla Nowego Outlooka nie wyliczamy lokalnej zajętości jako profilu pocztowego, aby nie przedstawiać cache aplikacji jako danych skrzynki."],
            )
        return MailProgramResult(self.OUTLOOK_NEW, False, "Nie wykryto")
