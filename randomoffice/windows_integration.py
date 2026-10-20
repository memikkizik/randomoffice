"""Read-only Windows installation state and explicit default-app settings access.

Registration is performed by ``scripts/install_desktop.ps1``. Importing this
module, inspecting status, and building a plan never write to the registry.
The app registers as a candidate; only Windows Settings chooses the defaults.
"""

from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path, PureWindowsPath
import sys
from typing import Any
from urllib.parse import quote


APP_NAME = "RandomOffice"
CAPABILITIES_KEY = r"Software\RandomOffice\Capabilities"
REGISTERED_APPLICATIONS_KEY = r"Software\RegisteredApplications"
FILE_TYPES = (
    (".docx", "RandomOffice.Document", "RandomOffice Word Document"),
    (".xlsx", "RandomOffice.Workbook", "RandomOffice Excel Workbook"),
    (".pptx", "RandomOffice.Presentation", "RandomOffice PowerPoint Presentation"),
)


@dataclass(frozen=True)
class RegistryValue:
    """One REG_SZ value under HKEY_CURRENT_USER; an empty name is (Default)."""

    key: str
    name: str
    value: str


def _validated_executable(executable_path: str | Path) -> str:
    raw = str(executable_path)
    if any(character in raw for character in ('"', "\r", "\n", "\0")):
        raise ValueError("Executable path contains an invalid character")
    path = PureWindowsPath(raw)
    if not path.is_absolute() or path.name.lower() != "randomoffice.exe":
        raise ValueError("An absolute Windows path to RandomOffice.exe is required")
    if ".." in path.parts:
        raise ValueError("Executable path must not contain parent-directory traversal")
    return str(path)


def build_registry_plan(executable_path: str | Path) -> tuple[RegistryValue, ...]:
    """Describe per-user integration without changing existing default choices.

    Extension keys receive only an OpenWithProgids value. Their default values
    and Explorer's protected UserChoice keys are deliberately never assigned.
    """
    executable = _validated_executable(executable_path)
    command = f'"{executable}" "%1"'
    icon = f'"{executable}",0'
    application = r"Software\Classes\Applications\RandomOffice.exe"
    values = [
        RegistryValue(CAPABILITIES_KEY, "ApplicationName", APP_NAME),
        RegistryValue(
            CAPABILITIES_KEY,
            "ApplicationDescription",
            "Create, edit and save Word documents, Excel workbooks and PowerPoint presentations locally.",
        ),
        RegistryValue(CAPABILITIES_KEY, "ApplicationIcon", icon),
        RegistryValue(REGISTERED_APPLICATIONS_KEY, APP_NAME, CAPABILITIES_KEY),
        RegistryValue(application, "FriendlyAppName", APP_NAME),
        RegistryValue(application + r"\DefaultIcon", "", icon),
        RegistryValue(application + r"\shell\open\command", "", command),
        RegistryValue(r"Software\Microsoft\Windows\CurrentVersion\App Paths\RandomOffice.exe", "", executable),
    ]
    for extension, progid, description in FILE_TYPES:
        progid_key = "Software\\Classes\\" + progid
        values.extend(
            [
                RegistryValue(progid_key, "", description),
                RegistryValue(progid_key, "FriendlyTypeName", description),
                RegistryValue(progid_key + r"\DefaultIcon", "", icon),
                RegistryValue(progid_key + r"\shell\open\command", "", command),
                RegistryValue("Software\\Classes\\" + extension + r"\OpenWithProgids", progid, ""),
                RegistryValue(application + r"\SupportedTypes", extension, ""),
                RegistryValue(CAPABILITIES_KEY + r"\FileAssociations", extension, progid),
            ]
        )
    return tuple(values)


def _is_windows() -> bool:
    return sys.platform == "win32"


def get_install_directory() -> Path:
    """The stable per-user directory shared by the installer and application."""
    local_app_data = os.environ.get("LOCALAPPDATA")
    base = Path(local_app_data) if local_app_data else Path.home() / "AppData" / "Local"
    return base / APP_NAME


def _read_registry_value(key: str, name: str) -> str | None:
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as handle:
            value, _ = winreg.QueryValueEx(handle, name)
            return value if isinstance(value, str) else None
    except (ImportError, OSError):
        return None


def default_apps_settings_uri() -> str:
    """Use the app-specific Windows 11 page, or the Windows 10 defaults page."""
    try:
        if _is_windows() and sys.getwindowsversion().build >= 22000:
            return "ms-settings:defaultapps?registeredAppUser=" + quote(APP_NAME, safe="")
    except AttributeError:
        pass
    return "ms-settings:defaultapps"


def get_install_status() -> dict[str, Any]:
    """Return installation and registration status without changing the system."""
    directory = get_install_directory()
    executable = directory / "RandomOffice.exe"
    supported = _is_windows()
    installed = supported and executable.is_file()
    registered = False
    if installed:
        # Validate every advertised association, including the actual launch
        # commands, so stale registrations to a moved executable are reported.
        plan = build_registry_plan(executable)
        registered = all(_read_registry_value(value.key, value.name) == value.value for value in plan)
    return {
        "supported": supported,
        "installed": installed,
        "registered": registered,
        "executable_path": str(executable),
        "install_directory": str(directory),
        "settings_uri": default_apps_settings_uri(),
        "extensions": [extension for extension, _, _ in FILE_TYPES],
    }


def open_default_apps_settings() -> dict[str, Any]:
    """Open Windows' own chooser; never programmatically force a default app."""
    if not _is_windows():
        return {"success": False, "error": "Varsayılan uygulama ayarları yalnızca Windows üzerinde kullanılabilir."}
    uri = default_apps_settings_uri()
    try:
        os.startfile(uri)
    except (AttributeError, OSError) as exc:
        return {"success": False, "error": f"Windows ayarları açılamadı: {exc}"}
    return {"success": True, "uri": uri}
