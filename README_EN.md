# xlsxDiff GUI 1.2.1 — English guide

Windows GUI derivative of [xlsxDiff](https://github.com/rafal-dot/xlsxDiff)
by **Rafał Czeczótka**, original copyright © 2020–2023.
2026 GUI changes maintained by **Michał Dubrawski**, with implementation and
documentation assistance from **OpenAI ChatGPT (GPT-6)**. Licensed under
**GNU AGPL-3.0-or-later** (see LICENSE and NOTICE.txt).

## Usage

1. Select your **older** `.xlsx` or `.xlsm` file and the **newer** `.xlsx` or `.xlsm` file (mixed input pairs are supported).
2. Choose an output file. Add the pair to the queue; you can add more pairs.
3. Change comparison options and click **Start comparison**.
4. Red strikethrough text means removed; blue underlined text means added.
   Whole-number differences can be shown as complete old → new values.
5. Sheet tabs: white indicates changed, blue added, red removed, gray usually
   unchanged (check gray tabs as well due to original detection limitations).

VBA macros are **not compared, copied, or executed**; the DIFF report is always a macro-free `.xlsx` file.

The report is not a full-fidelity copy of the newer workbook, and formulas
are not recalculated. When comparing values, cached spreadsheet values are used.

The **Polski / English** dropdown changes the GUI language instantly without
losing queued file pairs or selected options. It cannot be changed while a
comparison is running. The choice resets to the build default on next launch.

## Create the Windows executables (CMD)

Python 3.12–3.14 is expected; use the version available on your system.

```bat
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-build.txt
.venv\Scripts\python.exe -m unittest discover -s tests -v
build_onefile.cmd en
build_onefile.cmd pl
```

`dist\XlsxDiff_GUI_EN.exe` starts in English;
`dist\XlsxDiff_GUI_PL.exe` starts in Polish. **Both can switch languages.**
`build_onedir.cmd en` / `build_onedir.cmd pl` build directory distributions.
The build scripts embed `assets\xlsxDiff.ico` and the PNG for the app window.
You need to run these builds on **Windows** (this source package contains no
prebuilt Windows binary).

Direct Python start:

```bat
.venv\Scripts\python.exe xlsxDiff_GUI.py --lang en
```

Classic CLI remains available:

```bat
.venv\Scripts\python.exe xlsxDiff.py old.xlsm new.xlsx DIFF.xlsx -N
```

`-N` selects whole-number replacement, enabled by default only in GUI.

## Source release / legal notices

The planned URL (`https://github.com/mdubi/xlsxDiff-GUI`) is set in
`project_info.py` (`SOURCE_REPOSITORY_URL`). **Before publishing an EXE, create
that public repository, confirm the link works**, and rebuild. This activates
**About → Source code**. See `RELEASE_CHECKLIST.md` and `SOURCE.md`.

Publish the full corresponding source for the **matching binary version**,
retain the original author/copyright and prominent modification notices,
provide the AGPL license and relevant bundled third-party license notices,
and never upload proprietary customer spreadsheets to a public repository.

To generate a reviewed source-only release archive, run `package_source.cmd`
(or `python package_source.py`). Check the ZIP before publishing to GitHub.

The original xlsxDiff is AGPL-3.0-or-later. This is a modified derivative,
not the original author's official GUI. No warranty. These notes are
practical guidance, not legal advice.

## Repository / releases

1. Create a new **public** GitHub repository named `xlsxDiff-GUI` under `mdubi`
   (do not initialize it with a README, .gitignore or license).
2. From this directory, run:

   ```bat
   git init
   git branch -M main
   git add .
   git commit -m "Release xlsxDiff GUI 1.2.1 (AGPL-3.0-or-later)"
   git remote add origin https://github.com/mdubi/xlsxDiff-GUI.git
   git push -u origin main
   ```

3. Review repository contents before tagging; build and test the Windows
   binaries **from the same commit**, then create and push tag `v1.2.1`:

   ```bat
   git tag v1.2.1
   git push origin v1.2.1
   ```

4. Create a GitHub Release for `v1.2.1` and attach the two binaries, matching
   `XlsxDiff_GUI_v1.2.1_SOURCE.zip`, build manifest and required notices.
   GitHub provides automatic source archives for the exact tag as well.

5. Before public binary release, audit licenses of **all bundled dependencies**,
   including Python/Tcl/Tk, as described in `THIRD_PARTY.md`.

Do not include customer files, employee names in data, passwords, certificates,
real survey spreadsheets, or local build folders. Run `git status` and
`git ls-files` before pushing. This is general licensing guidance, not legal advice.
