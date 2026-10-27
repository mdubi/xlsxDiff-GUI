# Publication audit — 2026-10-27

## Verified in the source tree

- [x] Original `upstream_original/xlsxDiff.py` byte-for-byte equals the script
  in the user-supplied upstream archive (`xlsxDiff-main.zip`).
- [x] `LICENSE` byte-for-byte equals the upstream AGPL v3 license.
- [x] Public test fixtures (`tests/in1.xlsx`, `tests/in2.xlsx`) byte-for-byte
  equal the upstream test inputs, not customer files.
- [x] Clear author credit and 2026 modification dates appear in README,
  NOTICE, code and GUI About dialog.
- [x] The About dialog shows AGPL permissions, absence of warranty and
  access to the license text and source link.
- [x] English README and English user/build guide, Polish guide,
  English release notes, multi-size app icon, and paired build scripts.
- [x] 18 Python regression tests pass with a virtual X display on Linux,
  including two actual Tk GUI tests and XLSX generation.
- [x] Reviewed source-only archive follows an explicit file allowlist.

## Still required for actual public distribution

- [ ] Create and publish the repository `mdubi/xlsxDiff-GUI` (currently only
  a proposed URL in `project_info.py`).
- [ ] Upload the reviewed sources and verify the public GitHub page.
- [ ] Test both PyInstaller Windows executable flavors in a clean Windows VM
  or PC, especially Python 3.14.x, Polish and English, Excel output and icon.
- [ ] Record exact dependency/build versions and resulting EXE hashes.
- [ ] Collect and review all legally required third-party binary notices;
  bundle these notices with each binary release.
- [ ] Tag the exact source commit (`v1.2.1`) and attach corresponding source
  and any tested EXEs to the same GitHub Release.
- [ ] Keep customer spreadsheets, secrets, and private paths out of repository.

**Do not advertise a downloadable EXE or completed public release until all
applicable unchecked items are satisfied.** This audit provides practical
technical checks; it does not replace professional legal advice.
