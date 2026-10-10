# Issue #36 — v1.1.0 release preparation evidence

## Provenance and scope

- Repository: `bokoboss/tmc-processor`
- Issue: https://github.com/bokoboss/tmc-processor/issues/36
- Accepted baseline: `main@203bc5f3f584d0245f46c27d58da74c3d973fd65`
- Branch: `codex/issue-36-release-v1.1.0`
- Scope: package version alignment, Windows Support QR packaging regression,
  focused tests, stale release/deployment documentation, and candidate release
  bundles.
- Protected behavior: no calculation, Mapping, QC, Peak, V1/V2 export logic,
  Excel template, drawing, or chart source was changed.

## Validation

All commands ran on Windows with the repository `.venv` Python 3.12.10.

| Gate | Command / method | Result |
|---|---|---|
| Packaging/version regressions | `python -m pytest --basetemp .pytest-tmp-issue36 -xvv tests/test_windows_release_bundle.py` | 30 passed; deterministic runtime bundle, allowlist, private-file rejection, QR byte equality, missing-QR failure, extracted runtime paths, version label and app banner covered |
| Support QR/dialog | `python -m pytest --basetemp .pytest-tmp-issue36 -q tests/test_support_view.py -k "support_qr or support_dialog"` | 5 passed, 3 deselected |
| V1 report version | `python -m pytest --basetemp .pytest-tmp-issue36 -xvv tests/test_metadata_export.py::test_safe_png_workbook_exports_setup_metadata_to_metadata_and_setup_sheets` | 1 passed; `Export_Metadata.app_version` equals `APP_VERSION` |
| V1 OOXML export | `python -m pytest --basetemp .pytest-tmp-issue36 -xvv tests/test_ooxml_template_export.py` | 31 passed, 2 existing OpenPyXL DrawingML warnings |
| V2 export and round-once totals | `python -m pytest --basetemp .pytest-tmp-issue36 -xvv tests/test_v2_pcu_rounding.py` | 17 passed, 9 existing OpenPyXL DrawingML warnings; version metadata, formulas/caches, native drawings/charts and accepted demo totals covered |
| Synthetic demo smoke | `python scripts/smoke_demo.py` | PASS; 4 raw sheets, 2,304 normalized rows, AM 08:00 / PM 17:00, workbook generated |
| Compile | `python -m compileall -q app.py src scripts` | PASS |
| Whitespace | `git diff --check` | PASS; Git emitted only its informational LF-to-CRLF notice for the edited V2 test file |
| Candidate ZIP build | `python scripts/build_windows_release.py --demo --output dist/issue36-v1.1.0-candidate` | PASS; both ZIPs and manifests generated |
| ZIP integrity/manifest | independent CRC, member list, per-file byte count/SHA, archive SHA, prohibited-content, version and extraction checks | PASS for both archives |

The complete `test_support_view.py` run was started but manually stopped after
its third test remained active for over a minute. The focused Support QR/dialog
tests passed; the repository PR CI remains a required full-suite gate.

## QR provenance

- Accepted baseline tracked size: `1,130,446` bytes.
- Working-tree size: `1,130,446` bytes.
- SHA-256: `07cb6abdc58680ecf333ec4295f9c5ec9a1dd66b68268647ebbc267b02d52e31`.
- The regression compares the archive payload to `git show HEAD:assets/support_qr.png`.

## Candidate bundle evidence

These are local, unpublished candidate files. Rebuild both bundles from the exact
approved post-merge release commit before publishing assets.

| Candidate asset | Files | Compressed bytes | SHA-256 |
|---|---:|---:|---|
| `TMC-Processor-v1.1.0-Windows.zip` | 63 | 1,215,722 | `35a2149102a41da875a13384c97a9fa9e2ea0734bcc343e27850ff3047144f9a` |
| `TMC-Processor-v1.1.0-Demo-Files.zip` | 8 | 55,551 | `0c99505e6a866f7ed66fb51b4041fa97e3421b96174644c1f33ba3c144c351d6` |

Candidate ZIPs and JSON manifests are under the ignored local path
`dist/issue36-v1.1.0-candidate/`. The Windows ZIP contains the launcher, both
V1/V2 template-map pairs, and the byte-identical Support QR. The Demo ZIP
contains only the allowlisted synthetic `samples/demo/` files.

## Changed files

- `pyproject.toml`
- `scripts/build_windows_release.py`
- `tests/test_windows_release_bundle.py`
- `tests/test_metadata_export.py`
- `tests/test_v2_pcu_rounding.py`
- `README.md`
- `PROJECT_PROFILE.md`
- `docs/releases/v1.1.0.md`
- `docs/development/EVIDENCE_PACKAGE_ISSUE36.md`

## Pending gates and limitations

- PR creation and all required GitHub CI jobs are pending at this evidence
  checkpoint.
- The live Streamlit app remains on the accepted current `main` version until
  merge; the v1.1.0 Cloud banner/export smoke and operator Support/session check
  require the separately approved post-merge deployment.
- Excel COM GUI recalculation is not independently observable in these checks.
- Batch Project Session remains unsupported because its current format does
  not round-trip Batch uploads and Batch Mapping Presets.
- No merge, tag, or release publication was performed.
