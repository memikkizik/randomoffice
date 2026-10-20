from pathlib import Path
from types import SimpleNamespace
import sys

import pytest

from randomoffice import windows_integration as integration


EXECUTABLE = r"C:\Users\Örnek Kullanıcı\AppData\Local\RandomOffice\RandomOffice.exe"


def test_registry_plan_quotes_paths_and_registers_each_supported_type():
    plan = integration.build_registry_plan(EXECUTABLE)
    values = {(entry.key, entry.name): entry.value for entry in plan}
    assert values[(integration.REGISTERED_APPLICATIONS_KEY, "RandomOffice")] == integration.CAPABILITIES_KEY
    assert values[(integration.CAPABILITIES_KEY, "ApplicationDescription")]
    for extension, progid, _ in integration.FILE_TYPES:
        assert values[(integration.CAPABILITIES_KEY + r"\FileAssociations", extension)] == progid
        assert values[("Software\\Classes\\" + progid + r"\shell\open\command", "")] == f'"{EXECUTABLE}" "%1"'
        assert values[("Software\\Classes\\" + extension + r"\OpenWithProgids", progid)] == ""


def test_registry_plan_never_claims_or_overwrites_current_defaults():
    plan = integration.build_registry_plan(EXECUTABLE)
    assert all("userchoice" not in entry.key.lower() for entry in plan)
    for extension, _, _ in integration.FILE_TYPES:
        # Adding an Open With candidate is permitted. Assigning the extension's
        # default ProgID, MIME or perceived type would change another app's data.
        assert all(entry.key.lower() != ("Software\\Classes\\" + extension).lower() for entry in plan)
    assert all(entry.key.startswith("Software\\") for entry in plan)
    assert len({(entry.key.lower(), entry.name.lower()) for entry in plan}) == len(plan)


@pytest.mark.parametrize(
    "path",
    [
        "RandomOffice.exe",
        r"C:RandomOffice.exe",
        r"C:\Office\..\RandomOffice.exe",
        r'C:\Office\RandomOffice.exe" --extra',
        "C:\\Office\\RandomOffice.exe\n",
        r"C:\Office\other.exe",
    ],
)
def test_registry_plan_rejects_ambiguous_launch_paths(path):
    with pytest.raises(ValueError):
        integration.build_registry_plan(path)


def test_registry_reader_uses_current_user_and_closes_handle(monkeypatch):
    events = []

    class Key:
        def __enter__(self):
            return "opened-key"

        def __exit__(self, *args):
            events.append("closed")

    def open_key(root, key):
        assert root == "HKCU"
        assert key == "Software\\RandomOffice"
        return Key()

    def query_value(key, name):
        assert (key, name) == ("opened-key", "ApplicationName")
        return "RandomOffice", 1

    monkeypatch.setitem(sys.modules, "winreg", SimpleNamespace(HKEY_CURRENT_USER="HKCU", OpenKey=open_key, QueryValueEx=query_value))
    assert integration._read_registry_value("Software\\RandomOffice", "ApplicationName") == "RandomOffice"
    assert events == ["closed"]


def test_registry_reader_treats_missing_or_denied_key_as_unregistered(monkeypatch):
    def denied(*args):
        raise PermissionError("not available")

    monkeypatch.setitem(sys.modules, "winreg", SimpleNamespace(HKEY_CURRENT_USER="HKCU", OpenKey=denied))
    assert integration._read_registry_value("Software\\RandomOffice", "ApplicationName") is None


def test_install_status_detects_stale_launch_command(monkeypatch):
    directory = Path(r"C:\Users\Test\AppData\Local\RandomOffice")
    monkeypatch.setattr(integration, "get_install_directory", lambda: directory)
    monkeypatch.setattr(integration, "_is_windows", lambda: True)
    monkeypatch.setattr(Path, "is_file", lambda self: True)
    plan = integration.build_registry_plan(directory / "RandomOffice.exe")
    registry = {(entry.key, entry.name): entry.value for entry in plan}
    monkeypatch.setattr(integration, "_read_registry_value", lambda key, name: registry.get((key, name)))
    assert integration.get_install_status()["registered"] is True
    registry[(r"Software\Classes\RandomOffice.Document\shell\open\command", "")] = r'"C:\Temp\RandomOffice.exe" "%1"'
    status = integration.get_install_status()
    assert status["installed"] is True
    assert status["registered"] is False
    assert status["extensions"] == [".docx", ".xlsx", ".pptx"]


def test_missing_install_does_not_read_registry(monkeypatch):
    monkeypatch.setattr(integration, "_is_windows", lambda: True)
    monkeypatch.setattr(Path, "is_file", lambda self: False)

    def unexpected_read(*args):
        raise AssertionError("Missing executable must not be reported as installed")

    monkeypatch.setattr(integration, "_read_registry_value", unexpected_read)
    status = integration.get_install_status()
    assert status["installed"] is False
    assert status["registered"] is False


def test_per_user_install_directory_uses_local_app_data(monkeypatch, tmp_path):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    assert integration.get_install_directory() == tmp_path / "RandomOffice"


@pytest.mark.parametrize("build,expected", [(19045, "ms-settings:defaultapps"), (22631, "ms-settings:defaultapps?registeredAppUser=RandomOffice")])
def test_default_app_settings_uses_supported_windows_page(monkeypatch, build, expected):
    monkeypatch.setattr(integration, "_is_windows", lambda: True)
    monkeypatch.setattr(sys, "getwindowsversion", lambda: SimpleNamespace(build=build), raising=False)
    assert integration.default_apps_settings_uri() == expected


def test_default_app_settings_only_launches_windows_chooser(monkeypatch):
    launched = []
    monkeypatch.setattr(integration, "_is_windows", lambda: True)
    monkeypatch.setattr(integration, "default_apps_settings_uri", lambda: "ms-settings:defaultapps")
    monkeypatch.setattr(integration.os, "startfile", launched.append, raising=False)
    assert integration.open_default_apps_settings() == {"success": True, "uri": "ms-settings:defaultapps"}
    assert launched == ["ms-settings:defaultapps"]


def test_default_app_settings_reports_launch_failure(monkeypatch):
    monkeypatch.setattr(integration, "_is_windows", lambda: True)

    def failed(uri):
        raise OSError("Settings unavailable")

    monkeypatch.setattr(integration.os, "startfile", failed, raising=False)
    result = integration.open_default_apps_settings()
    assert result["success"] is False
    assert "Settings unavailable" in result["error"]


def test_default_app_settings_reports_unsupported_platform(monkeypatch):
    monkeypatch.setattr(integration, "_is_windows", lambda: False)
    assert integration.get_install_status()["supported"] is False
    assert integration.open_default_apps_settings()["success"] is False
