from __future__ import annotations

import importlib.util
from io import BytesIO
from pathlib import Path
import shutil
import subprocess
import zipfile
from xml.etree import ElementTree

import pytest


ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("windows_release", ROOT / "scripts/build_windows_release.py")
release = importlib.util.module_from_spec(spec)
spec.loader.exec_module(release)


def test_runtime_bundle_is_complete_deterministic_and_private_files_stay_out(tmp_path):
    source = tmp_path / "source"
    for name in release.RUNTIME_FILES:
        target = source / name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(ROOT / name, target)
    shutil.copytree(ROOT / "src/tmc_processor", source / "src/tmc_processor", ignore=shutil.ignore_patterns("__pycache__"))
    for name in (".streamlit/secrets.toml", "samples/raw/private.xlsx", "outputs/private.xlsx",
                 "src/tmc_processor/__pycache__/private.pyc", "tests/private.py", "AGENTS.md",
                 ".git/config", ".github/workflows/ci.yml", ".venv/private.py", "docs/evidence.xlsx"):
        target = source / name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(b"PRIVATE_SENTINEL")
    first = release.build(source, tmp_path / "first")[0]
    # File timestamps must have no effect on output.
    (source / "app.py").touch()
    second = release.build(source, tmp_path / "second")[0]
    assert first.read_bytes() == second.read_bytes()
    with zipfile.ZipFile(first) as archive:
        assert archive.testzip() is None
        prefix = f"{first.stem}/"
        names = {n.removeprefix(prefix) for n in archive.namelist()}
        expected_sources = {p.relative_to(source).as_posix() for p in (source / "src/tmc_processor").rglob("*.py")}
        assert names == set(release.RUNTIME_FILES) | expected_sources
        assert all(b"PRIVATE_SENTINEL" not in archive.read(n) for n in archive.namelist())


@pytest.mark.parametrize("path", [
    ".git/config", ".github/workflows/ci.yml", ".venv/python.exe", "tests/test.py",
    "samples/raw/private.xlsx", "samples/demo/DEMO_TMC1_FourLeg.xlsx", "outputs/report.xlsx",
    "src/tmc_processor/__pycache__/x.pyc", "src/tmc_processor/tests/private.py",
    "AGENTS.md", "PROJECT_PROFILE.md", "docs/evidence.md", ".streamlit/secrets.toml",
    "../app.py", "/app.py", "C:\\app.py", "src/tmc_processor/x.egg-info/private.py",
])
def test_prohibited_paths_are_rejected(path):
    with pytest.raises(ValueError):
        release.validate_path(path)


def test_missing_required_template_fails_before_build(tmp_path):
    source = tmp_path / "source"
    source.mkdir()
    with pytest.raises((FileNotFoundError, ValueError)):
        release.build(source, tmp_path / "dist")
    assert not (tmp_path / "dist").exists()


@pytest.mark.parametrize("data", [b'C:\\MyRD\\private.xlsx', b'C:/Users/operator/private.xlsx', b'file:///private.xlsx'])
def test_embedded_local_path_is_rejected(data):
    with pytest.raises(ValueError):
        release.validate_content("app.py", data)


def test_embedded_ooxml_local_path_is_rejected(tmp_path):
    workbook = tmp_path / "template.xlsx"
    with zipfile.ZipFile(workbook, "w") as archive:
        archive.writestr("xl/_rels/workbook.xml.rels", '<Relationship Target="C:/MyRD/private.xlsx"/>')
    with pytest.raises(ValueError):
        release.validate_content("template.xlsx", workbook.read_bytes())


@pytest.mark.parametrize("name", ["four_leg_tmc_report_template.xlsx", "four_leg_tmc_report_template_approach_v2.xlsx"])
def test_portable_template_preserves_every_part_except_saved_location(name):
    original = (ROOT / "templates" / name).read_bytes()
    portable = release.portable_template(original)
    assert release.portable_template(portable) == portable
    release.validate_content(name, portable)
    with zipfile.ZipFile(BytesIO(original)) as before, zipfile.ZipFile(BytesIO(portable)) as after:
        assert before.namelist() == after.namelist()
        for part in before.namelist():
            expected = before.read(part)
            if part == "xl/workbook.xml":
                expected, count = release.SAVED_LOCATION.subn(b"", expected)
                assert count == 1
                ElementTree.fromstring(expected)
            assert after.read(part) == expected


def test_demo_uses_only_allowlisted_committed_bytes(monkeypatch, tmp_path):
    requested = []
    def committed_bytes(command):
        requested.append(command)
        return (ROOT / command[-1].removeprefix("HEAD:")).read_bytes()
    monkeypatch.setattr(subprocess, "check_output", committed_bytes)
    (tmp_path / "samples/raw").mkdir(parents=True)
    (tmp_path / "samples/raw/private.xlsx").write_bytes(b"PRIVATE_SENTINEL")
    for name in release.DEMO_FILES:
        local = tmp_path / name
        local.parent.mkdir(parents=True, exist_ok=True)
        local.write_bytes(b"LOCAL_MODIFICATION_SENTINEL")
    payload = release.demo_payload(tmp_path)
    assert set(payload) == set(release.DEMO_FILES)
    assert all(data == (ROOT / name).read_bytes() for name, data in payload.items())
    assert [c[-1] for c in requested] == [f"HEAD:{name}" for name in release.DEMO_FILES]
    with pytest.raises(ValueError):
        release.validate_path("samples/raw/private.xlsx", demo=True)
