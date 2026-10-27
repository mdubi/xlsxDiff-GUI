"""Localization, icon, launch configuration and optional GUI smoke tests."""
import os
import sys
import tempfile
import time
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from PIL import Image
from i18n import STRINGS, tr
from xlsxDiff_GUI import APP_VERSION, App, suggest_output


class LocalizationTests(unittest.TestCase):
    def test_all_translations_have_both_languages(self):
        for key, pair in STRINGS.items():
            self.assertEqual(len(pair), 2, key)
            for lang in ('pl', 'en'):
                self.assertTrue(tr(lang, key, count=2, file="diff.xlsx", cells=2, sheets=1, added=0, removed=0).strip(), (lang, key))
        self.assertEqual(tr('en', 'col_old'), 'OLDER FILE')
        self.assertEqual(tr('pl', 'col_old'), 'STARA WERSJA')

    def test_icon_contains_windows_sizes(self):
        with Image.open(ROOT / 'assets/xlsxDiff.ico') as im:
            self.assertTrue({(16,16),(32,32),(48,48),(256,256)} <= im.ico.sizes())
        self.assertTrue((ROOT / 'assets/xlsxDiff_icon_512.png').is_file())
        self.assertEqual(APP_VERSION, '1.2.1')

    def test_both_launchers_use_correct_default(self):
        self.assertIn("default_language='pl'", (ROOT/'launcher_pl.py').read_text(encoding='utf8'))
        self.assertIn("default_language='en'", (ROOT/'launcher_en.py').read_text(encoding='utf8'))
        for filename in ('build_onefile.cmd','build_onedir.cmd'):
            script=(ROOT/filename).read_text(encoding='utf8').lower()
            self.assertIn('--icon', script)
            self.assertIn('assets\\xlsxdiff.ico', script)
            self.assertIn('launcher_en.py', script)
            self.assertIn('launcher_pl.py', script)
            self.assertIn('license', script)


@unittest.skipUnless(sys.platform == 'win32' or os.environ.get('DISPLAY'),
                     'Requires an actual Tk desktop or virtual DISPLAY')
class GuiTests(unittest.TestCase):
    def setUp(self):
        self.app = App('pl')
        self.app.withdraw()
        self.app.update_idletasks()

    def tearDown(self):
        if self.app.winfo_exists():
            # Tk keeps timer callbacks across short-lived test windows.
            # Cancel them before destroying the interpreter to avoid callbacks
            # targeting a destroyed instance during a later test.
            for callback_id in self.app.tk.call('after', 'info'):
                self.app.after_cancel(callback_id)
            self.app.destroy()

    def test_switch_language_preserves_queue_and_options(self):
        old = ROOT/'tests/in1.xlsx'
        new = ROOT/'tests/in2.xlsx'
        with tempfile.TemporaryDirectory() as d:
            self.app.old_var.set(str(old))
            self.app.new_var.set(str(new))
            self.app.out_var.set(str(Path(d)/'diff.xlsx'))
            self.app._add_job()
            self.app.columns_var.set('staff!A')
            self.app.rows_var.set('staff!1')
            self.app.whole_numbers.set(False)
            self.app.display_lang.set('English')
            self.app._change_language()
            self.app.update_idletasks()
            self.assertEqual(self.app.lang, 'en')
            self.assertEqual(self.app.tree.heading('old')['text'], 'OLDER FILE')
            self.assertEqual(self.app.tree.item('1','values')[-1], 'Queued')
            self.assertEqual(self.app.columns_var.get(), 'staff!A')
            self.assertEqual(self.app.rows_var.get(), 'staff!1')
            self.assertFalse(self.app.whole_numbers.get())
            self.app.display_lang.set('Polski')
            self.app._change_language()
            self.assertEqual(self.app.tree.item('1','values')[-1], 'Oczekuje')
            self.assertEqual(self.app.tree.get_children(), ('1',))

    def test_gui_english_end_to_end(self):
        with tempfile.TemporaryDirectory() as d:
            self.app.display_lang.set('English')
            self.app._change_language()
            output = Path(d)/'DIFF.xlsx'
            # End-to-end GUI test uses two XLSM inputs, exercising the new
            # accepted extension through the actual comparison worker.
            from openpyxl import Workbook
            files = (Path(d)/'old.xlsm', Path(d)/'new.xlsm')
            for fn, value in zip(files, ('Before', 'After')):
                wb = Workbook()
                wb.active['A1'] = value
                wb.save(fn)
                wb.close()
            self.app.old_var.set(str(files[0]))
            self.app.new_var.set(str(files[1]))
            self.app.out_var.set(str(output))
            self.app._add_job()
            self.app._start()
            deadline = time.monotonic() + 8
            while self.app.is_running and time.monotonic() < deadline:
                self.app.update()
                time.sleep(.03)
            self.assertFalse(self.app.is_running)
            self.assertTrue(output.is_file())
            self.assertEqual(self.app.jobs['1'].status, 'done')
            self.assertEqual(self.app.tree.item('1','values')[-1], 'Done')


if __name__ == '__main__':
    unittest.main()
