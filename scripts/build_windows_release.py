"""Build a lean Windows source bundle offline; Python is required to run it."""

from __future__ import annotations

import argparse
import hashlib
from io import BytesIO
import json
from pathlib import Path, PurePosixPath
import re
import subprocess
import zipfile

try:
    import tomllib
except ModuleNotFoundError:  # Python 3.10, already a project dependency
    import tomli as tomllib


ROOT = Path(__file__).resolve().parents[1]
# Editable installation needs metadata plus source. Both template/map pairs are
# referenced by report_template.py. Only the public theme config is included.
RUNTIME_FILES = (
    "LICENSE", "USER_GUIDE.md", "app.py", "pyproject.toml",
    "start_tmc_processor.bat", ".streamlit/config.toml",
    "assets/support_qr.png",
    "templates/four_leg_tmc_report_template.xlsx",
    "templates/four_leg_tmc_report_template_map.json",
    "templates/four_leg_tmc_report_template_approach_v2.xlsx",
    "templates/four_leg_tmc_report_template_approach_v2_map.json",
)
# These committed fixtures are synthetic; --demo reads HEAD, never local files.
DEMO_FILES = (
    "samples/demo/DEMO_TMC1_FourLeg.mapping.json",
    "samples/demo/DEMO_TMC1_FourLeg.xlsx",
    "samples/demo/DEMO_TMC1_FourLeg_Day2.xlsx",
    "samples/demo/DEMO_TMC1_FourLeg_approach_v2.mapping.json",
    "samples/demo/DEMO_TMC1_FourLeg_approach_v2_mapping.xlsx",
    "samples/demo/DEMO_TMC1_FourLeg_mapping.xlsx",
    "samples/demo/DEMO_TMC1_FourLeg_session.tmcproj.json",
    "samples/demo/README.md",
)
PROHIBITED = {
    ".git", ".github", ".venv", "venv", ".pytest_cache", "__pycache__",
    "tests", "docs", "outputs", "raw", "legacy_output", "dist", "build",
    "agents.md", "project_profile.md", ".engineering-workflow.json",
    ".engineering-workflow", ".env", "secrets.toml",
}
LOCAL_PATH = re.compile(rb"(?i)(?<![a-z])(?:[a-z]:[\\/]|file:///|/users/|/home/|/myrd/|\\\\[^\\\s]+\\)")
# Only Excel's saved-location metadata in the two supplied templates is removed.
# MS-XLSX 2.4.67: absPath; no cells, relationships or other XML are rewritten.
SAVED_LOCATION = re.compile(
    rb'<mc:AlternateContent\b[^>]*><mc:Choice Requires="x15">'
    rb'<x15ac:absPath url="[^"]*" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/>'
    rb'</mc:Choice></mc:AlternateContent>'
)


def portable_template(data: bytes) -> bytes:
    """Remove saved-location provenance from bundled copies, preserving all other parts."""
    with zipfile.ZipFile(BytesIO(data)) as source:
        original = source.read("xl/workbook.xml")
        portable, count = SAVED_LOCATION.subn(b"", original)
        if not count:
            return data
        target = BytesIO()
        with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
            for info in source.infolist():
                content = portable if info.filename == "xl/workbook.xml" else source.read(info.filename)
                archive.writestr(info, content, compresslevel=9)
        return target.getvalue()


def validate_path(name: str, *, demo: bool = False) -> None:
    path = PurePosixPath(name)
    if path.is_absolute() or "\\" in name or ".." in path.parts:
        raise ValueError(f"Unsafe archive path: {name}")
    if any(part.lower() in PROHIBITED or part.lower().endswith(".egg-info") for part in path.parts):
        raise ValueError(f"Prohibited archive path: {name}")
    if demo:
        allowed = name in DEMO_FILES
    else:
        allowed = name in RUNTIME_FILES or (
            name.startswith("src/tmc_processor/") and path.suffix == ".py"
            and "samples" not in path.parts
        )
    if not allowed:
        raise ValueError(f"Path outside allowlist: {name}")


def validate_content(name: str, data: bytes) -> None:
    # Inspect embedded OOXML parts as well as ordinary text. A workbook ZIP is
    # not opaque: external relationships can otherwise conceal local paths.
    parts = [(name, data)]
    if name.endswith(".xlsx"):
        with zipfile.ZipFile(BytesIO(data)) as workbook:
            parts = [(f"{name}:{part}", workbook.read(part)) for part in workbook.namelist()]
    for part, content in parts:
        if LOCAL_PATH.search(content) or LOCAL_PATH.search(content.replace(b"\x00", b"")):
            raise ValueError(f"Local absolute path embedded in {part}")


def runtime_payload(root: Path) -> dict[str, bytes]:
    root = root.resolve()
    names = list(RUNTIME_FILES)
    names.extend(p.relative_to(root).as_posix() for p in (root / "src/tmc_processor").rglob("*.py"))
    if "src/tmc_processor/__init__.py" not in names:
        raise ValueError("Package source is missing")
    payload = {}
    for name in sorted(names):
        validate_path(name)
        file = root / name
        if file.is_symlink() or not file.resolve().is_relative_to(root):
            raise ValueError(f"Source escapes repository: {name}")
        payload[name] = file.read_bytes()
        if name.startswith("templates/") and name.endswith(".xlsx"):
            payload[name] = portable_template(payload[name])
        validate_content(name, payload[name])
    return payload


def demo_payload(root: Path) -> dict[str, bytes]:
    payload = {}
    for name in DEMO_FILES:
        validate_path(name, demo=True)
        payload[name] = subprocess.check_output(["git", "-C", str(root), "show", f"HEAD:{name}"])
        validate_content(name, payload[name])
    return payload


def write_bundle(payload: dict[str, bytes], output: Path, stem: str, *, demo: bool = False) -> Path:
    output.mkdir(parents=True, exist_ok=True)
    archive_path = output / f"{stem}.zip"
    # Fixed timestamps, permissions, order and compression give repeatable ZIPs.
    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for name, data in sorted(payload.items()):
            validate_path(name, demo=demo)
            info = zipfile.ZipInfo(f"{stem}/{name}", date_time=(1980, 1, 1, 0, 0, 0))
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, data, compresslevel=9)
    with zipfile.ZipFile(archive_path) as archive:
        if archive.testzip() is not None:
            raise ValueError("ZIP integrity check failed")
        if archive.namelist() != [f"{stem}/{name}" for name in sorted(payload)]:
            raise ValueError("ZIP manifest mismatch")
    manifest = {
        "archive": archive_path.name,
        "sha256": hashlib.sha256(archive_path.read_bytes()).hexdigest(),
        "compressed_bytes": archive_path.stat().st_size,
        "extracted_bytes": sum(map(len, payload.values())),
        "files": [{"path": f"{stem}/{name}", "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()}
                  for name, data in sorted(payload.items())],
    }
    (output / f"{stem}.manifest.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(manifest, indent=2))
    return archive_path


def build(root: Path = ROOT, output: Path | None = None, *, demo: bool = False) -> list[Path]:
    payload = runtime_payload(root)
    version = tomllib.loads(payload["pyproject.toml"].decode("utf-8"))["project"]["version"]
    if not re.fullmatch(r"\d+\.\d+\.\d+(?:[-+][A-Za-z0-9.-]+)?", version):
        raise ValueError("Unsafe or unsupported release version")
    output = output or root / "dist"
    # Validate all inputs before writing either requested artifact.
    demos = demo_payload(root) if demo else None
    result = [write_bundle(payload, output, f"TMC-Processor-v{version}-Windows")]
    if demos is not None:
        result.append(write_bundle(demos, output, f"TMC-Processor-v{version}-Demo-Files", demo=True))
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, help="Output directory (default: dist)")
    parser.add_argument("--demo", action="store_true", help="Also package committed synthetic demos (requires Git)")
    args = parser.parse_args()
    build(output=args.output, demo=args.demo)
