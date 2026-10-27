# xlsxDiff GUI v1.2.1 — XLSM input support

This release adds XLSM comparison input support to the 1.2.0 GUI and CLI.
The tool now accepts `.xlsx` and `.xlsm` input workbooks in any pairing.
The generated diff report remains a plain `.xlsx` file; VBA macros are not
compared, copied, executed or embedded in the report. The original input
files remain unchanged.

All other comparison behavior and options are unchanged.

Original xlsxDiff: © 2020–2023 Rafał Czeczótka (AGPL-3.0-or-later).
GUI changes: © 2026 Michał Dubrawski, assisted by ChatGPT (GPT-6).
See README.md, LICENSE and NOTICE.txt before redistribution.

Build Windows executables from this exact source revision using
`build_onefile.cmd en` and `build_onefile.cmd pl` and test before release.
