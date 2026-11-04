"""Lesson 61: readable summaries and structured CLI errors."""

import json

from src import recovery_cli


def _args(tmp_path, output=None):
    return [
        "--excel", str(tmp_path / "missing.xlsx"),
        "--audit", str(tmp_path / "missing.json"),
        "--output", str(output or tmp_path / "recovery.json"),
    ]


def test_text_summary_creates_json_report(tmp_path, capsys):
    output = tmp_path / "recovery.json"
    code = recovery_cli.main(_args(tmp_path) + ["--format", "text"])
    captured = capsys.readouterr()
    assert code == 2
    assert "Recovery status : NOT_PUBLISHED" in captured.out
    assert "Safe to use     : NO" in captured.out
    assert "Saved report" in captured.out
    assert json.loads(output.read_text(encoding="utf-8"))["safe_to_use"] is False


def test_default_display_remains_json(tmp_path, capsys):
    code = recovery_cli.main(_args(tmp_path))
    assert code == 2
    data = json.loads(capsys.readouterr().out)
    assert data["recovery_status"] == "NOT_PUBLISHED"


def test_structured_errors_never_overwrite_report(tmp_path, capsys):
    report = tmp_path / "recovery.json"
    report.write_text("existing content", encoding="utf-8")
    code = recovery_cli.main(_args(tmp_path) + ["--error-format", "json"])
    captured = capsys.readouterr()
    assert code == 1
    assert captured.out == ""
    assert json.loads(captured.err)["report_created"] is False
    assert report.read_text(encoding="utf-8") == "existing content"


def test_text_errors_remain_default(tmp_path, capsys):
    report = tmp_path / "recovery.json"
    report.write_text("keep", encoding="utf-8")
    assert recovery_cli.main(_args(tmp_path)) == 1
    assert "Recovery report not created:" in capsys.readouterr().err


def test_verified_text_summary(monkeypatch, capsys):
    monkeypatch.setattr(recovery_cli, "create_recovery_report", lambda **kw: {
        "recovery_status": "VERIFIED", "safe_to_use": True,
        "manual_review_required": False, "recommended_action": "No recovery needed.",
    })
    code = recovery_cli.main([
        "--excel", "sample.xlsx", "--audit", "audit.json", "--output", "new.json",
        "--format", "text",
    ])
    assert code == 0
    assert "Safe to use     : YES" in capsys.readouterr().out


def test_invalid_format_cannot_create_report(tmp_path):
    import pytest
    with pytest.raises(SystemExit) as error:
        recovery_cli.main(_args(tmp_path) + ["--format", "xml"])
    assert error.value.code == 2
    assert not (tmp_path / "recovery.json").exists()


def test_text_report_preserves_input_file(tmp_path):
    excel = tmp_path / "synthetic.xlsx"
    excel.write_bytes(b"not an excel file")
    audit = tmp_path / "audit.json"
    audit.write_text("not-json", encoding="utf-8")
    original_excel = excel.read_bytes()
    original_audit = audit.read_bytes()
    output = tmp_path / "assessment.json"
    code = recovery_cli.main([
        "--excel", str(excel), "--audit", str(audit),
        "--output", str(output), "--format", "text",
    ])
    assert code == 2
    assert excel.read_bytes() == original_excel
    assert audit.read_bytes() == original_audit
    assert output.exists()
