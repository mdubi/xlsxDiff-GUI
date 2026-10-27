# xlsxDiff GUI v1.2.0 — first public source release

An independent bilingual Windows GUI adaptation of
[Rafał Czeczótka's xlsxDiff](https://github.com/rafal-dot/xlsxDiff),
maintained by **Michał Dubrawski** with development assistance from OpenAI
ChatGPT (GPT-6). Available under **GNU AGPL-3.0-or-later**.

### Highlights

- Compare two XLSX workbooks, worksheet by worksheet and cell by cell.
- Highlight deleted content in red with strikethrough and inserted content in
  blue with underline.
- Optional full replacement diff for numbers and percentages.
- Multiple-pair comparison queue, progress status, cancellation and logging.
- English and Polish GUI, switchable while the app is idle.
- Windows icon and optional English-default / Polish-default builds.
- Original CLI remains available.

### Release files

- `XlsxDiff_GUI_v1.2.0_SOURCE.zip`: full matching source and build scripts.
- `XlsxDiff_GUI_EN.exe` and `XlsxDiff_GUI_PL.exe`: **add only after both have
  been built and tested on Windows**. Both contain both languages; suffix
  specifies the initial language.
- Build manifest and third-party license texts: **add once audited**.

**No executables have been published as of the preparation of these notes.**
Do not claim an executable was tested or attached until confirmed.

### Important limitations

- This is a DIFF report, not an editable copy of the newer workbook.
- No formula calculation; saved/cached values are compared by default.
- Always check gray tab contents for edge cases and validate critical changes.

### Attribution and license

Original xlsxDiff © 2020–2023 Rafał Czeczótka.
2026 modifications maintained by Michał Dubrawski with assistance from
OpenAI ChatGPT (GPT-6). This project is not an official upstream release.
See `LICENSE`, `NOTICE.txt`, `SOURCE.md` and `THIRD_PARTY.md`.
The code is offered under GNU AGPL-3.0-or-later, without warranty.
