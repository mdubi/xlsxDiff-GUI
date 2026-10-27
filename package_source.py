# -*- coding: utf-8 -*-
"""Create a safe, complete *application* source snapshot for GitHub Releases.

Only the explicitly reviewed public source files and upstream test fixtures are
included. No client spreadsheets, secrets, local paths, venv, or build output.
"""
from __future__ import annotations

import argparse
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parent
VERSION = '1.2.1'
PUBLIC_FILES = (
    '.gitignore', 'README.md', 'README_PL.md', 'README_EN.md',
    'LICENSE', 'NOTICE.txt', 'SOURCE.md', 'THIRD_PARTY.md',
    'RELEASE_CHECKLIST.md', 'CHANGELOG.md',
    'PUBLICATION_AUDIT.md', 'RELEASE_NOTES_v1.2.0.md', 'RELEASE_NOTES_v1.2.1.md',
    'requirements.txt', 'requirements-build.txt',
    'build_onefile.cmd', 'build_onedir.cmd', 'run_gui.cmd',
    'package_source.py', 'package_source.cmd',
    'xlsxDiff.py', 'xlsxDiff_GUI.py', 'xlsx_diff_engine.py',
    'i18n.py', 'project_info.py', 'launcher_pl.py', 'launcher_en.py',
    'assets/xlsxDiff.ico', 'assets/xlsxDiff_icon_master.png',
    'assets/xlsxDiff_icon_512.png',
    'upstream_original/README.md', 'upstream_original/xlsxDiff.py',
    'tests/in1.xlsx', 'tests/in2.xlsx',
    'tests/test_engine.py', 'tests/test_gui_i18n.py',
)


def create_source_package(destination: Path) -> Path:
    missing = [item for item in PUBLIC_FILES if not (ROOT / item).is_file()]
    if missing:
        raise FileNotFoundError('Missing required source files: ' + ', '.join(missing))
    destination = destination.resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(destination, 'w', compression=ZIP_DEFLATED, compresslevel=9) as zf:
        for relative in sorted(PUBLIC_FILES):
            zf.write(ROOT / relative, arcname=f'XlsxDiff_GUI/{relative}')
    return destination


def main(argv=None):
    parser = argparse.ArgumentParser(description='Create a reviewable source archive (no private files).')
    parser.add_argument('destination', nargs='?', default=f'XlsxDiff_GUI_v{VERSION}_SOURCE.zip')
    args = parser.parse_args(argv)
    print('Source archive:', create_source_package(Path(args.destination)))
    print('Review the ZIP contents before publishing, particularly files in tests/.')


if __name__ == '__main__':
    main()
