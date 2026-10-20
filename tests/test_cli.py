from randomoffice.__main__ import main


def test_cli_reports_backend_status(capsys) -> None:
    assert main([]) == 0

    output = capsys.readouterr().out
    assert "RandomOffice backend is ready." in output
    assert ".docx, .xlsx, .pptx" in output
