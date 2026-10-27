# Third-party dependencies and notices

The main **modified xlsxDiff work** is licensed under GNU AGPL-3.0-or-later.
The following independently maintained projects may be present in the packaged
Windows executable or needed to build/test it.

| Component | Purpose | License / source |
|---|---|---|
| CPython standard library | Runtime | [PSF License](https://docs.python.org/3/license.html) |
| Tcl/Tk | Graphical interface | [Tcl/Tk licensing](https://www.tcl.tk/software/tcltk/license.html) |
| openpyxl | Read XLSX | MIT, [source](https://foss.heptapod.net/openpyxl/openpyxl) |
| XlsxWriter | Write XLSX report | BSD 2-Clause, [source and license](https://xlsxwriter.readthedocs.io/license.html) |
| et_xmlfile | Dependency of openpyxl | Verify exact installed package metadata and license |
| PyInstaller | Generate executable | GPL-2.0-or-later with special bootloader exception, [project](https://pyinstaller.org/) |
| Pillow | Test/build icon support (not normally required at app runtime) | HPND, [project](https://python-pillow.github.io/pillow/) |
| Other transitively bundled packages | Packaging/runtime | Check the actual PyInstaller dependency collection |

**Important:** this is a guide, **not** a complete collection of notices for
Windows binaries. Before distributing any EXE, inspect its exact bundled
components (including Python, Tcl/Tk and PyInstaller's bootloader), collect
any required copyright/license notices **into the release materials**, and
preserve conditions applying to binary distributions. For example, XlsxWriter
BSD-2-Clause requires copyright/license/disclaimer notices with binaries.

### Preparing a binary release

1. Build from a fresh `.venv` on Windows, using documented versions.
2. Record `python --version` and a `pip freeze` from **that build venv**.
3. Inspect the final PyInstaller bundle for actual included dependencies.
4. Copy applicable exact-version license texts into a separate
   `THIRD_PARTY_LICENSES/` folder accompanying the release, including PSF and
   Tcl/Tk terms. Verify the complete set yourself, rather than assuming a
   Python package's license is inherited from its dependencies.
5. Publish these notices and sufficiently complete corresponding sources/build
   instructions with the corresponding EXE. Review AGPL §6 and project terms.

No Windows binaries or exact-version binary license bundle are included in
this source-only package. Verify again for every new binary build.
