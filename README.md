# xlsxDiff GUI

**A bilingual Windows desktop app for comparing Excel `.xlsx` and `.xlsm` workbooks** with
Word-like tracked changes inside cells. Inspired by and derived from
[Rafał Czeczótka’s original xlsxDiff](https://github.com/rafal-dot/xlsxDiff).

**Version:** 1.2.1 · **Languages:** English / Polish · **License:**
[GNU AGPL-3.0-or-later](LICENSE)

[English user & build guide](README_EN.md) · [Instrukcja po polsku](README_PL.md) ·
[Source & license obligations](SOURCE.md) · [Release checklist](RELEASE_CHECKLIST.md) ·
[Publication audit](PUBLICATION_AUDIT.md)

## What it does

- Accepts XLSX and XLSM inputs (including mixed pairs), compares worksheet cells,
  ignores VBA, and always writes a macro-free XLSX report.
- Matches worksheets by name and compares corresponding cells.
- Shows removed text in **red strikethrough** and inserted text in **blue underline**.
- Optionally treats a changed number / percentage as **one replacement** rather
  than coloring only individual changed digits (enabled by default in the GUI).
- Supports formula comparison, row/column alignment indexes, header highlighting,
  and other options inherited from the command-line program.
- Provides a queue of workbook pairs, progress display, cancellation and a log.
- Can switch the full graphical interface between **English and Polish** without
  restarting. Windows builds may start in either language.
- Ships with a multi-resolution Windows icon (`assets/xlsxDiff.ico`).

**Limitations:** the output is a *new comparison report*, not a faithful copy
of the newer workbook's formatting. Formula values depend on what was last
cached in the original Excel files; this app does not calculate formulas.
Worksheet colors are indicators, not a substitute for checking cells—older
xlsxDiff sheet-change detection has edge cases. Always validate critical edits.

## Quick start on Windows

Python 3.12–3.14 and a command prompt in the project directory:

```bat
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-build.txt
.venv\Scripts\python.exe -m unittest discover -s tests -v
.venv\Scripts\python.exe xlsxDiff_GUI.py --lang en
```

Build an executable (Windows required):

```bat
build_onefile.cmd en
build_onefile.cmd pl
```

The outputs are `dist\XlsxDiff_GUI_EN.exe` and
`dist\XlsxDiff_GUI_PL.exe`. They differ **only in startup language**;
each supports the runtime language switch. For folder distributions instead,
run `build_onedir.cmd en` or `build_onedir.cmd pl`.

The command-line interface is also available:

```bat
.venv\Scripts\python.exe xlsxDiff.py older.xlsx newer.xlsx DIFF.xlsx -N
```

`-N` selects whole-number replacement in CLI mode (GUI enables it by default).
See the [English guide](README_EN.md) for all options and examples.

## Credits and provenance

- **Original xlsxDiff:** [Rafał Czeczótka](https://github.com/rafal-dot/xlsxDiff),
  © 2020–2023, distributed under GNU AGPL-3.0-or-later.
- **2026 GUI adaptation and maintenance:** **Michał Dubrawski**.
- **Development assistance:** OpenAI **ChatGPT (GPT-6)** contributed to
  implementation, tests, interface design and documentation. The human
  maintainer is responsible for the published changes. OpenAI is not a
  project maintainer, licensor or endorser.

This repository is an **independent modified derivative**, not an official
release from the original author. Original code is retained unchanged in
[`upstream_original/xlsxDiff.py`](upstream_original/xlsxDiff.py); detailed
modification notices appear in [`NOTICE.txt`](NOTICE.txt) and the
[change log](CHANGELOG.md).

## License and distribution

The modified project is provided under **AGPL-3.0-or-later**, without warranty.
You may use, study, modify and redistribute it under the license terms.
If distributing a Windows executable, provide the **Corresponding Source for
that exact build**, retain copyright/modification notices, and account for
third-party licenses and the build environment. See [SOURCE.md](SOURCE.md),
[THIRD_PARTY.md](THIRD_PARTY.md) and [RELEASE_CHECKLIST.md](RELEASE_CHECKLIST.md).

The repository is intended to be published at
`https://github.com/mdubi/xlsxDiff-GUI`. **Until that repository is public,
this is only a planned address; no public release or executable download is
being claimed.** Do not redistribute release builds with a dead source link.

## Contributing and reporting issues

You are welcome to open an issue or pull request after the public repository is
created. Please do **not** upload client spreadsheets or other confidential
workbooks to GitHub. Use small synthetic XLSX/XLSM files to demonstrate problems.
