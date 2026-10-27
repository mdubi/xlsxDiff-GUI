# Release and AGPL checklist / Wydanie i licencja

This is a practical checklist, not legal advice.

- [ ] License: `LICENSE` (GNU AGPL-3.0-or-later) included, not replaced by MIT.
- [ ] Credit the original author **Rafał Czeczótka**, © 2020–2023, and link
      https://github.com/rafal-dot/xlsxDiff in README, NOTICE and About window.
- [ ] Keep the original source in `upstream_original/xlsxDiff.py` and retain
      all original copyright/license notices.
- [ ] `NOTICE.txt` prominently documents that the code was modified and when,
      and credits Michał Dubrawski as the maintainer of 2026 modifications.
- [ ] README and NOTICE disclose ChatGPT (GPT-6) development assistance
      without presenting OpenAI as project owner, endorser, or license holder.
- [ ] Release our derivative GUI/engine source under AGPL-3.0-or-later as a whole.
- [ ] Create the proposed **public** repo `mdubi/xlsxDiff-GUI`, then verify
      the `project_info.py` URL actually resolves **before building binaries**.
- [ ] Publish full source at exactly the same Git tag as each binary release
      (include tests, icons, launchers, build scripts, complete instructions).
- [ ] Run `package_source.cmd`; examine its ZIP contents for private information.
- [ ] Attach a corresponding source ZIP to each GitHub Release (or provide
      another compliant source distribution mechanism) **alongside** any EXE.
- [ ] Include `LICENSE`, `NOTICE.txt`, `SOURCE.md`, `THIRD_PARTY.md`
      with downloadable binaries, and prominent notices inside GUI About dialog.
- [ ] Document exact versions of Python, dependencies and PyInstaller;
      include relevant bundled third-party license texts/source as required.
- [ ] Test the compiled PL and EN builds on Windows, including the About and
      Source-code dialogs and Windows icon/taskbar visibility.
- [ ] Do not include customers' XLSX files, real reports, credentials, network
      drives, `.venv`, build artifacts, caches, personal logs or certificates
      in the public source repository.
- [ ] If network-based use of a future version is introduced, review AGPL §13.

A public GitHub repository is **recommended but not inherently required** by
AGPL. Distributors must still comply with source and license obligations.

Also consult `PUBLICATION_AUDIT.md` and draft `RELEASE_NOTES_v1.2.1.md`.
