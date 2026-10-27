# Changelog

## 1.2.1 — 2026-10-27
- Accept XLSM files as comparison inputs in CLI and bilingual GUI, including
  mixed XLSX/XLSM pairs. VBA is never executed, copied or compared.
- Preserve XLSX-only output and add regression tests for XLSM input.

## 1.2.0 — 2026-10-27
- Publication preparation: English-first README and guide, named 2026 maintainer,
  AI-assistance disclosure, prominent AGPL permissions, public-source metadata.
- New Windows icon in multiple sizes, embedded in the EXE and GUI window.
- Fully bilingual PL/EN menus, dialogs, options and interface strings.
- Live language switching without losing the comparison queue/options.
- Two executable defaults: `build_onefile.cmd pl` and `build_onefile.cmd en`.
- Expanded credits and access-to-source action in About dialog.
- Release/AGPL guidance, bilingual documentation and GUI smoke tests.

## 1.1.0 — 2026-10-21
- Optional whole-number (including text percentages) difference mode.
- Credits for original xlsxDiff author and license notices in GUI.

## 1.0.0 — 2026-10-21
- Split upstream script into reusable comparison engine and CLI wrapper.
- Windows GUI with queue, status, cancel, logging and safer output paths.
- PyInstaller build scripts and regression tests.
