"""Tests for the Lesson 60 command-line interface."""

import json
import subprocess
import sys
from pathlib import Path

import pytest

from src import recovery_cli


def test_cli_not_published_creates_report(tmp_path, capsys):
    output = tmp_path / "assessment.json"
    code = recovery_cli.main([
        "--excel", str(tmp_path / "missing.xlsx"),
        "--audit", str(tmp_path / "missing.json"),
        "--output", str(output),
    ])
    assert code == 2
    data = json.loads(output.read_text(encoding="utf-8"))
    assert data["recovery_status"] == "NOT_PUBLISHED"
    assert data["safe_to_use"] is False
    assert "NOT_PUBLISHED" in capsys.readouterr().out


def test_cli_does_not_overwrite_existing_report(tmp_path, capsys):
    output = tmp_path / "assessment.json"
    output.write_text("original", encoding="utf-8")
    code = recovery_cli.main([
        "--excel", str(tmp_path / "missing.xlsx"),
        "--audit", str(tmp_path / "missing.json"),
        "--output", str(output),
    ])
    assert code == 1
    assert output.read_text(encoding="utf-8") == "original"
    assert "already exists" in capsys.readouterr().err


def test_cli_rejects_output_same_as_audit(tmp_path, capsys):
    audit = tmp_path / "audit.json"
    audit.write_text("keep", encoding="utf-8")
    code = recovery_cli.main([
        "--excel", str(tmp_path / "missing.xlsx"),
        "--audit", str(audit),
        "--output", str(audit),
    ])
    assert code == 1
    assert audit.read_text(encoding="utf-8") == "keep"


def test_cli_verified_path(monkeypatch, capsys):
    def fake_report(**kwargs):
        return {
            "recovery_status": "VERIFIED", "safe_to_use": True,
            "manual_review_required": False,
            "recommended_action": "No recovery needed.",
        }
    monkeypatch.setattr(recovery_cli, "create_recovery_report", fake_report)
    code = recovery_cli.main([
        "--excel", "example.xlsx", "--audit", "audit.json", "--output", "new.json",
    ])
    assert code == 0
    assert '"safe_to_use": true' in capsys.readouterr().out


def test_cli_requires_arguments():
    with pytest.raises(SystemExit) as exc:
        recovery_cli.main([])
    assert exc.value.code == 2


def test_module_help_runs_from_project_root():
    root = Path(__file__).resolve().parents[1]
    proc = subprocess.run(
        [sys.executable, "-m", "src.recovery_cli", "--help"],
        cwd=root, text=True, capture_output=True, check=False,
    )
    assert proc.returncode == 0
    assert "--excel" in proc.stdout
