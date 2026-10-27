# -*- coding: utf-8 -*-
"""Bilingual xlsxDiff desktop GUI; derivative of Rafał Czeczótka's xlsxDiff.

Copyright (C) 2020-2023 Rafał Czeczótka (original xlsxDiff)
2026 modifications maintained by Michał Dubrawski, with ChatGPT (GPT-6) assistance.
GNU AGPL-3.0-or-later; see LICENSE/NOTICE.txt.
"""
from __future__ import annotations

import argparse
import os
import queue
import re
import subprocess
import sys
import threading
import traceback
import webbrowser
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from i18n import tr
from project_info import SOURCE_REPOSITORY_URL
from xlsx_diff_engine import ComparisonCancelled, DiffOptions, compare_workbooks, parse_index_spec

APP_VERSION = '1.2.1'
UPSTREAM_URL = 'https://github.com/rafal-dot/xlsxDiff'
_BG = '#f4f7fa'
_NAVY = '#173452'


def resource_path(relative: str) -> Path:
    return Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent)) / relative


def _open_path(path: str):
    if sys.platform == 'win32':
        os.startfile(path)
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', path])
    else:
        subprocess.Popen(['xdg-open', path])


def suggest_output(older: str, newer: str) -> str:
    old, new = Path(older), Path(newer)
    name = f'DIFF_{old.stem[:26]}_vs_{new.stem[-32:]}.xlsx'
    candidate = new.with_name(name)
    n = 2
    while candidate.exists():
        candidate = new.with_name(f'{name[:-5]}_{n}.xlsx')
        n += 1
    return str(candidate)


@dataclass
class Job:
    older: str
    newer: str
    output: str
    status: str = 'queued'


class App(tk.Tk):
    def __init__(self, default_language='pl'):
        super().__init__()
        if default_language not in ('pl', 'en'):
            raise ValueError('Unsupported language')
        self.lang = default_language
        self.geometry('1100x805')
        self.minsize(880, 710)
        self.configure(background=_BG)
        self.jobs: dict[str, Job] = {}
        self.job_number = 0
        self.pending = queue.Queue()
        self.worker: threading.Thread | None = None
        self.cancel_event = threading.Event()
        self.is_running = False
        self._closing = False
        self.old_var = tk.StringVar()
        self.new_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.columns_var = tk.StringVar()
        self.rows_var = tk.StringVar()
        self.formulas = tk.BooleanVar(value=False)
        self.headers = tk.BooleanVar(value=False)
        self.skip_empty = tk.BooleanVar(value=False)
        self.autofilter = tk.BooleanVar(value=False)
        self.added_format = tk.BooleanVar(value=True)
        self.whole_numbers = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value=tr(self.lang, 'ready'))
        self.display_lang = tk.StringVar(value='English' if self.lang == 'en' else 'Polski')
        self.log_lines: list[tuple[str, str | None]] = []
        self._setup_style()
        self._set_icon()
        self._build()
        self.after(110, self._poll)
        self.protocol('WM_DELETE_WINDOW', self._on_close)

    def t(self, key, **kwargs):
        return tr(self.lang, key, **kwargs)

    def _set_icon(self):
        # Windows: use the multi-resolution ICO consistently.  Calling
        # iconphoto() afterwards would replace iconbitmap() on Windows
        # and let Tk scale the old full-size PNG for the taskbar.
        ico = resource_path('assets/xlsxDiff.ico')
        if sys.platform == 'win32' and ico.exists():
            try:
                self.iconbitmap(default=str(ico))
                return
            except tk.TclError:
                pass
        # Cross-platform fallback for systems that do not accept ICO.
        png = resource_path('assets/xlsxDiff_icon_512.png')
        if png.exists():
            try:
                self._icon_image = tk.PhotoImage(file=str(png))
                self.iconphoto(True, self._icon_image)
            except tk.TclError:
                pass

    def _setup_style(self):
        st = ttk.Style(self)
        st.theme_use('vista' if 'vista' in st.theme_names() else 'clam')
        st.configure('App.TFrame', background=_BG)
        st.configure('App.TLabel', background=_BG, foreground=_NAVY, font=('Segoe UI', 10))
        st.configure('Caption.TLabel', background=_BG, foreground='#526477', font=('Segoe UI', 9))
        st.configure('Title.TLabel', background=_BG, foreground=_NAVY, font=('Segoe UI Semibold', 21))
        st.configure('Sub.TLabel', background=_BG, foreground=_NAVY, font=('Segoe UI Semibold', 11))
        st.configure('Treeview', font=('Segoe UI', 9), rowheight=23)
        st.configure('Treeview.Heading', font=('Segoe UI Semibold', 9))
        st.configure('Main.Horizontal.TProgressbar', thickness=13)
        st.map('Treeview', background=[('selected', '#d6eafb')], foreground=[('selected', '#102b42')])

    def _change_language(self, _event=None):
        if self.is_running:
            return
        language = 'en' if self.display_lang.get() == 'English' else 'pl'
        if language != self.lang:
            self.lang = language
            self.status_var.set(self.t('ready'))
            self._build()

    def _build(self):
        if hasattr(self, 'container'):
            self.container.destroy()
        self.title(f'xlsxDiff GUI {APP_VERSION} — {self.t("title")}')
        self.container = ttk.Frame(self, padding=12, style='App.TFrame')
        self.container.pack(fill='both', expand=True)
        top = ttk.Frame(self.container, style='App.TFrame')
        top.pack(fill='x')
        ttk.Label(top, text=self.t('title'), style='Title.TLabel').pack(side='left', anchor='w')
        self.language_box = ttk.Combobox(top, state='readonly', textvariable=self.display_lang,
                                         width=9, values=['Polski', 'English'])
        self.language_box.pack(side='right', padx=(8, 0))
        self.language_box.bind('<<ComboboxSelected>>', self._change_language)
        ttk.Label(top, text=self.t('language'), style='Caption.TLabel').pack(side='right')
        ttk.Button(top, text=self.t('about_btn'), command=self._about).pack(side='right', padx=(0, 12))
        intro = ttk.Frame(self.container, style='App.TFrame')
        intro.pack(fill='x', pady=(2, 1))
        ttk.Label(intro, text=self.t('subtitle'), style='Caption.TLabel').pack(side='left')
        ttk.Label(intro, text=self.t('credits'), style='Caption.TLabel').pack(side='right')
        ttk.Label(self.container, textvariable=self.status_var, style='Caption.TLabel',
                  wraplength=960).pack(anchor='w', pady=(0, 5))

        files = ttk.LabelFrame(self.container, text='  ' + self.t('files') + '  ', padding=(10, 7))
        files.pack(fill='x', pady=(0, 8))
        files.columnconfigure(1, weight=1)
        for row, (key, var, func) in enumerate([
            ('older', self.old_var, self._browse_old),
            ('newer', self.new_var, self._browse_new),
            ('report', self.out_var, self._browse_output),
        ]):
            ttk.Label(files, text=self.t(key), width=16).grid(row=row, column=0, sticky='w', pady=3)
            ttk.Entry(files, textvariable=var).grid(row=row, column=1, sticky='ew', padx=8, pady=3)
            ttk.Button(files, text=self.t('browse'), width=12, command=func).grid(
                row=row, column=2, padx=4, pady=3)
        ttk.Button(files, text=self.t('add'), command=self._add_job).grid(
            row=3, column=1, sticky='e', padx=8, pady=(3, 0))
        ttk.Button(files, text=self.t('swap'), command=self._swap).grid(
            row=3, column=2, sticky='e', padx=4, pady=(3, 0))

        queue_frame = ttk.LabelFrame(self.container, text='  ' + self.t('queue') + '  ', padding=(8, 6))
        queue_frame.pack(fill='both', expand=True, pady=(0, 8))
        table_frame = ttk.Frame(queue_frame)
        table_frame.pack(fill='both', expand=True)
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(table_frame, columns=('old', 'new', 'out', 'state'),
                                 show='headings', selectmode='extended', height=4)
        for col, label, width in [('old', 'col_old', 270), ('new', 'col_new', 270),
                                  ('out', 'col_out', 285), ('state', 'col_status', 110)]:
            self.tree.heading(col, text=self.t(label))
            self.tree.column(col, width=width, minwidth=95, stretch=(col != 'state'))
        self.tree.grid(row=0, column=0, sticky='nsew')
        yscroll = ttk.Scrollbar(table_frame, orient='vertical', command=self.tree.yview)
        yscroll.grid(row=0, column=1, sticky='ns')
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.bind('<Double-1>', self._double_click)
        for key, job in self.jobs.items():
            self.tree.insert('', 'end', iid=key, values=(Path(job.older).name, Path(job.newer).name,
                            Path(job.output).name, self.t(job.status)))
        controls = ttk.Frame(queue_frame)
        controls.pack(fill='x', pady=(8, 0))
        self.remove_btn = ttk.Button(controls, text=self.t('remove'), command=self._remove_selected)
        self.remove_btn.pack(side='left', padx=(0, 6))
        self.up_btn = ttk.Button(controls, text='↑', width=3, command=lambda: self._move(-1))
        self.up_btn.pack(side='left', padx=2)
        self.down_btn = ttk.Button(controls, text='↓', width=3, command=lambda: self._move(1))
        self.down_btn.pack(side='left', padx=2)
        ttk.Button(controls, text=self.t('open_report'), command=self._open_report).pack(side='right', padx=(5, 0))
        ttk.Button(controls, text=self.t('open_folder'), command=self._open_folder).pack(side='right')

        notebook = ttk.Notebook(self.container)
        notebook.pack(fill='x', pady=(0, 7))
        advanced = ttk.Frame(notebook, padding=(9, 8))
        notebook.add(advanced, text='  ' + self.t('options') + '  ')
        ttk.Checkbutton(advanced, text=self.t('whole_numbers'), variable=self.whole_numbers).pack(
            anchor='w', pady=(0, 2))
        ttk.Label(advanced, text=self.t('whole_numbers_help'), style='Caption.TLabel',
                  wraplength=980).pack(anchor='w', pady=(0, 5))
        options = ttk.Frame(advanced)
        options.pack(fill='x')
        for index, (key, var) in enumerate([
            ('formulas', self.formulas), ('headers', self.headers),
            ('skip_empty', self.skip_empty), ('added_format', self.added_format),
            ('autofilter', self.autofilter),
        ]):
            ttk.Checkbutton(options, text=self.t(key), variable=var).grid(
                row=index // 2, column=index % 2, sticky='w', padx=(0, 35), pady=2)
        alignment = ttk.Frame(notebook, padding=(9, 8))
        notebook.add(alignment, text='  ' + self.t('alignment') + '  ')
        indices = ttk.Frame(alignment)
        indices.pack(fill='x', pady=(4, 0))
        indices.columnconfigure(1, weight=1)
        ttk.Label(indices, text=self.t('index_cols')).grid(row=0, column=0, sticky='w')
        ttk.Entry(indices, textvariable=self.columns_var).grid(row=0, column=1, sticky='ew', padx=(10, 0), pady=2)
        ttk.Label(indices, text=self.t('index_rows')).grid(row=1, column=0, sticky='w')
        ttk.Entry(indices, textvariable=self.rows_var).grid(row=1, column=1, sticky='ew', padx=(10, 0), pady=2)
        ttk.Label(indices, text=self.t('index_example'), style='Caption.TLabel').grid(
            row=2, column=0, columnspan=2, sticky='w', pady=(4, 0))

        log_frame = ttk.Frame(notebook, padding=5)
        notebook.add(log_frame, text='  ' + self.t('log_tab') + '  ')
        self.log = scrolledtext.ScrolledText(log_frame, height=7, wrap='word', bg='#fcfdfe',
                                             fg='#2b3a47', relief='flat', borderwidth=0,
                                             font=('Consolas', 9), state='disabled')
        self.log.pack(fill='both', expand=True)
        for tag, color in [('warn', '#a55b00'), ('error', '#a61b1b'), ('success', '#1f7c4c')]:
            self.log.tag_configure(tag, foreground=color)
        if not self.log_lines:
            self.log_lines.append((self.t('ready_log'), None))
        for msg, tag in self.log_lines:
            self.log.insert('end', msg + '\n', tag)
        self.log.configure(state='disabled')

        actions = ttk.Frame(self.container, style='App.TFrame')
        actions.pack(fill='x', pady=(1, 6))
        self.start_btn = ttk.Button(actions, text=self.t('run'), command=self._start)
        self.start_btn.pack(side='left')
        self.stop_btn = ttk.Button(actions, text=self.t('cancel'), command=self._cancel, state='disabled')
        self.stop_btn.pack(side='left', padx=8)
        self.progressbar = ttk.Progressbar(actions, mode='determinate', maximum=100,
                                           style='Main.Horizontal.TProgressbar')
        self.progressbar.pack(side='left', fill='x', expand=True, padx=(10, 0))

    def _about(self):
        dialog = tk.Toplevel(self)
        dialog.title(self.t('about_title'))
        dialog.resizable(False, False)
        dialog.transient(self)
        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill='both', expand=True)
        for key, style, space in [
            (f'xlsxDiff GUI {APP_VERSION}', 'Sub.TLabel', (0, 4)),
            (self.t('based_on'), 'App.TLabel', (0, 7)),
            (self.t('original_author'), 'App.TLabel', (0, 1)),
            (self.t('maintainer'), 'App.TLabel', (0, 2)),
            ('© 2020–2023 Rafał Czeczótka', 'App.TLabel', (0, 8)),
            (self.t('modifications'), 'App.TLabel', (0, 8)),
            (self.t('license'), 'App.TLabel', (0, 2)),
            (self.t('user_rights'), 'App.TLabel', (0, 3)),
            (self.t('no_warranty'), 'Caption.TLabel', (0, 9)),
        ]:
            ttk.Label(frame, text=key, style=style, wraplength=670).pack(anchor='w', pady=space)
        links = ttk.Frame(frame)
        links.pack(fill='x')
        ttk.Button(links, text=self.t('upstream_btn'), command=lambda: webbrowser.open(UPSTREAM_URL)).pack(side='left')
        ttk.Button(links, text=self.t('source_btn'), command=self._source_info).pack(side='left', padx=7)
        ttk.Button(links, text=self.t('license_btn'), command=lambda: self._show_license(dialog)).pack(side='left')
        ttk.Button(links, text=self.t('close'), command=dialog.destroy).pack(side='right')
        dialog.protocol('WM_DELETE_WINDOW', dialog.destroy)
        dialog.wait_visibility()
        dialog.grab_set()
        dialog.focus_set()

    def _source_info(self):
        if SOURCE_REPOSITORY_URL:
            webbrowser.open(SOURCE_REPOSITORY_URL)
        else:
            messagebox.showinfo(self.t('source_title'), self.t('source_info'), parent=self)

    def _show_license(self, parent):
        try:
            contents = resource_path('LICENSE').read_text(encoding='utf-8')
        except OSError:
            messagebox.showinfo(self.t('license_btn'), self.t('license_missing'), parent=parent)
            return
        win = tk.Toplevel(parent)
        win.title('GNU AGPL v3')
        win.geometry('780x550')
        win.transient(parent)
        frame = ttk.Frame(win, padding=10)
        frame.pack(fill='both', expand=True)
        txt = scrolledtext.ScrolledText(frame, wrap='word', font=('Consolas', 9))
        txt.pack(fill='both', expand=True)
        txt.insert('1.0', contents)
        txt.configure(state='disabled')
        def close():
            win.grab_release()
            win.destroy()
            if parent.winfo_exists():
                parent.grab_set()
        ttk.Button(frame, text=self.t('close'), command=close).pack(anchor='e', pady=(6, 0))
        win.protocol('WM_DELETE_WINDOW', close)
        win.wait_visibility()
        win.grab_set()

    def _choose_file(self, var):
        initial = Path(var.get()).parent if var.get() else None
        file = filedialog.askopenfilename(parent=self, title=self.t('select_xlsx'),
                    initialdir=str(initial) if initial and initial.exists() else None,
                    filetypes=[(self.t('excel_file'), '*.xlsx *.xlsm'), (self.t('all_files'), '*.*')])
        if file:
            var.set(file)
        return bool(file)

    def _browse_old(self):
        self._choose_file(self.old_var)

    def _browse_new(self):
        if self._choose_file(self.new_var) and self.old_var.get():
            self.out_var.set(suggest_output(self.old_var.get(), self.new_var.get()))

    def _browse_output(self):
        current = self.out_var.get() or (suggest_output(self.old_var.get(), self.new_var.get())
                    if self.old_var.get() and self.new_var.get() else 'DIFF.xlsx')
        file = filedialog.asksaveasfilename(parent=self, title=self.t('save_report'),
                    initialdir=str(Path(current).parent), initialfile=Path(current).name,
                    defaultextension='.xlsx', filetypes=[(self.t('excel_file'), '*.xlsx')])
        if file:
            self.out_var.set(file)

    def _swap(self):
        o, n = self.old_var.get(), self.new_var.get()
        self.old_var.set(n)
        self.new_var.set(o)
        if n and o:
            self.out_var.set(suggest_output(n, o))

    def _add_job(self):
        if self.is_running:
            return
        old = self.old_var.get().strip().strip('"')
        new = self.new_var.get().strip().strip('"')
        if not old or not new:
            messagebox.showwarning(self.t('no_files_title'), self.t('no_files_msg'))
            return
        output = self.out_var.get().strip().strip('"') or suggest_output(old, new)
        try:
            if any(not Path(f).is_file() or Path(f).suffix.lower() not in ('.xlsx', '.xlsm') for f in (old, new)):
                raise ValueError(self.t('not_xlsx'))
            canonical = [Path(p).resolve() for p in (old, new, output)]
            if len(set(canonical)) != 3:
                raise ValueError(self.t('distinct'))
            if Path(output).suffix.lower() != '.xlsx':
                raise ValueError(self.t('output_xlsx'))
            if any(Path(job.output).resolve() == canonical[2] for job in self.jobs.values()):
                raise ValueError(self.t('duplicate_out'))
        except ValueError as err:
            messagebox.showerror(self.t('invalid_files_title'), str(err))
            return
        self.job_number += 1
        key = str(self.job_number)
        self.jobs[key] = Job(old, new, output)
        self.tree.insert('', 'end', iid=key,
                         values=(Path(old).name, Path(new).name, Path(output).name, self.t('queued')))
        self.tree.selection_set(key)
        self.tree.see(key)
        self.old_var.set('')
        self.new_var.set('')
        self.out_var.set('')
        self.status_var.set(self.t('queued_count', count=len(self.jobs)))

    def _remove_selected(self):
        if self.is_running:
            return
        for key in self.tree.selection():
            self.jobs.pop(key, None)
            self.tree.delete(key)

    def _move(self, offset):
        if self.is_running:
            return
        selected = self.tree.selection()
        if len(selected) != 1:
            return
        key = selected[0]
        index = self.tree.index(key)
        self.tree.move(key, '', max(0, min(len(self.tree.get_children()) - 1, index + offset)))
        # Keep model insertion order consistent when rebuilding after language change.
        order = self.tree.get_children()
        self.jobs = {k: self.jobs[k] for k in order}

    def _selected_job(self):
        selected = self.tree.selection()
        return self.jobs[selected[0]] if selected else None

    def _double_click(self, _event):
        if self.is_running:
            return
        job = self._selected_job()
        if job:
            self.old_var.set(job.older)
            self.new_var.set(job.newer)
            self.out_var.set(suggest_output(job.older, job.newer))

    def _open_report(self):
        job = self._selected_job()
        if not job:
            messagebox.showinfo(self.t('choose_report_title'), self.t('choose_report_msg'))
        elif Path(job.output).is_file():
            try:
                _open_path(job.output)
            except OSError as err:
                messagebox.showerror(self.t('opening_err_title'), str(err))
        else:
            messagebox.showinfo(self.t('no_report_title'), self.t('no_report_msg'))

    def _open_folder(self):
        job = self._selected_job()
        if job:
            folder = Path(job.output).parent
            if folder.is_dir():
                try:
                    _open_path(str(folder))
                except OSError as err:
                    messagebox.showerror(self.t('opening_err_title'), str(err))

    def _get_options(self):
        return DiffOptions(compare_formulas=self.formulas.get(),
            highlight_headers=self.headers.get(),
            index_columns=parse_index_spec(self.columns_var.get(), 'column'),
            index_rows=parse_index_spec(self.rows_var.get(), 'row'),
            highlight_added_removed=self.added_format.get(),
            autofilter=self.autofilter.get(), skip_empty=self.skip_empty.get(),
            whole_number_diff=self.whole_numbers.get())

    def _set_status(self, key, status):
        self.jobs[key].status = status
        if self.tree.exists(key):
            vals = list(self.tree.item(key, 'values'))
            vals[3] = self.t(status)
            self.tree.item(key, values=vals)

    def _start(self):
        if self.is_running:
            return
        if not self.jobs:
            messagebox.showwarning(self.t('empty_queue_title'), self.t('empty_queue_msg'))
            return
        try:
            opts = self._get_options()
            if opts.autofilter and not opts.highlight_headers:
                raise ValueError(self.t('autofilter_req'))
        except ValueError as err:
            messagebox.showerror(self.t('invalid_options'), str(err))
            return
        ids = self.tree.get_children()
        existing = [self.jobs[k].output for k in ids if Path(self.jobs[k].output).exists()]
        if existing and not messagebox.askyesno(self.t('overwrite_title'),
                    self.t('overwrite_msg', count=len(existing))):
            return
        self.cancel_event.clear()
        self.is_running = True
        self.progressbar['value'] = 0
        self.start_btn.configure(state='disabled')
        self.stop_btn.configure(state='normal')
        self.language_box.configure(state='disabled')
        for btn in (self.remove_btn, self.up_btn, self.down_btn):
            btn.configure(state='disabled')
        for key in ids:
            self._set_status(key, 'queued')
        self.worker = threading.Thread(target=self._run_jobs, args=(ids, opts), daemon=True)
        self.worker.start()

    def _run_jobs(self, ids, options):
        total = len(ids)
        for number, key in enumerate(ids):
            if self.cancel_event.is_set():
                break
            job = self.jobs[key]
            self.pending.put(('started', key, f'{number + 1}/{total}: {Path(job.newer).name}'))
            try:
                stats = compare_workbooks(job.older, job.newer, job.output, options,
                    progress=lambda fraction, msg, k=key, n=number: self.pending.put(
                        ('progress', k, (n + fraction) * 100 / total, msg)),
                    is_cancelled=self.cancel_event.is_set,
                    warn=lambda msg: self.pending.put(('warn', msg)))
                self.pending.put(('done', key, stats, job.output))
            except ComparisonCancelled:
                self.pending.put(('cancelled', key))
                break
            except Exception as err:
                self.pending.put(('error', key, str(err), traceback.format_exc()))
        self.pending.put(('finished',))

    def _localize_progress(self, msg):
        lookup = {'Wczytywanie starszej wersji…': 'status_load_old',
                  'Wczytywanie nowszej wersji…': 'status_load_new', 'Gotowe': 'status_done'}
        if msg in lookup:
            return self.t(lookup[msg])
        if self.lang == 'en':
            msg = re.sub(r'^Porównywanie arkusza ', 'Comparing sheet ', msg)
            msg = re.sub(r'^Arkusz ', 'Sheet ', msg)
        return msg

    def _poll(self):
        try:
            while True:
                event = self.pending.get_nowait()
                kind = event[0]
                if kind == 'started':
                    self._set_status(event[1], 'started')
                    self.status_var.set(self.t('starting', file=event[2]))
                    self._append_log('▶ ' + event[2])
                elif kind == 'progress':
                    self.progressbar['value'] = event[2]
                    self.status_var.set(self._localize_progress(event[3]))
                elif kind == 'done':
                    key, stats, path = event[1], event[2], event[3]
                    self._set_status(key, 'done')
                    self._append_log(self.t('stats', file=Path(path).name,
                        cells=stats.changed_cells, sheets=stats.changed_sheets,
                        added=stats.new_sheets, removed=stats.removed_sheets), 'success')
                elif kind == 'warn':
                    self._append_log('⚠ ' + event[1], 'warn')
                elif kind == 'error':
                    self._set_status(event[1], 'error')
                    self._append_log('✖ ' + event[2], 'error')
                    self._append_log(event[3], 'error')
                elif kind == 'cancelled':
                    self._set_status(event[1], 'cancelled')
                    self._append_log(self.t('cancelled_log'), 'warn')
                elif kind == 'finished':
                    self.is_running = False
                    self.start_btn.configure(state='normal')
                    self.stop_btn.configure(state='disabled')
                    self.language_box.configure(state='readonly')
                    for btn in (self.remove_btn, self.up_btn, self.down_btn):
                        btn.configure(state='normal')
                    if self.cancel_event.is_set():
                        self.status_var.set(self.t('interrupted'))
                    else:
                        self.progressbar['value'] = 100
                        self.status_var.set(self.t('finished'))
                    if self._closing:
                        self.destroy()
                        return
        except queue.Empty:
            pass
        if not self._closing or self.is_running:
            self.after(110, self._poll)

    def _cancel(self):
        if self.is_running:
            self.cancel_event.set()
            self.stop_btn.configure(state='disabled')
            self.status_var.set(self.t('cancelling'))

    def _append_log(self, msg, tag=None):
        self.log_lines.append((msg, tag))
        self.log.configure(state='normal')
        self.log.insert('end', msg + '\n', tag)
        self.log.see('end')
        self.log.configure(state='disabled')

    def _on_close(self):
        if self.is_running:
            if not messagebox.askyesno(self.t('close_running_title'), self.t('close_running_msg')):
                return
            self._closing = True
            self._cancel()
        else:
            self.destroy()


def main(default_language='pl', argv=None):
    parser = argparse.ArgumentParser(description='xlsxDiff GUI (PL/EN)')
    parser.add_argument('--lang', choices=['pl', 'en'], default=default_language,
                        help='GUI language; build determines default (pl or en)')
    args = parser.parse_args(argv)
    App(args.lang).mainloop()


if __name__ == '__main__':
    main()
