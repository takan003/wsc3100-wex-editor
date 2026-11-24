import tkinter as tk
from tkinter import ttk
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from wex_editor import EditorApp, ScheduleImportDialog
from wex_io import ScheduleRow


def schedule_row(music_id, name):
    return ScheduleRow("MUSIC1", {
        "MUSIC": music_id,
        "Music_no": music_id,
        "Music_name": name,
        "Clock_year": "2015",
        "Clock_month": "03",
        "Clock_day": "14",
        "Clock_start_hour": "08",
        "Clock_start_min": "00",
        "Mode": 1,
        "Music_on": 1,
    })


class FakeDatabase:
    def read_table(self, entry):
        return SimpleNamespace(columns=[SimpleNamespace(name="Music_name", col_size=100)])

    def read_columns(self, table):
        pass


class FakeProject:
    tab_names = {1: "First", 2: "Second", 3: "Empty"}
    audio_durations = {}
    audio_files = ["tone.wav"]
    group_names = {}
    db = FakeDatabase()
    _catalog = {name: object() for name in ("MUSIC", "MUSIC1", "MUSIC2", "MUSIC3", "MUSIC_T")}
    schedule_ids = [1, 2, 3]
    tabs = [
        (1, "First", [schedule_row(1, "First row")]),
        (2, "Second", [schedule_row(1, "Second row")]),
        (3, "Empty", []),
    ]

    def load_rows(self, table):
        if table == "MUSIC_T":
            return self.tabs[1][2]
        for sid, _, rows in self.tabs:
            if table == f"MUSIC{sid}":
                return rows
        return []

    def next_music_id(self, table):
        return max((row.music_id for row in self.load_rows(table)), default=0) + 1


class EditorTests(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.app = EditorApp(self.root)
        self.root.update()
        self.ws = self.app.active_workspace

    def tearDown(self):
        for job in self.root.tk.call("after", "info"):
            self.root.tk.call("after", "cancel", job)
        self.root.destroy()

    def load_project(self):
        with patch("wex_editor.WexProject", return_value=FakeProject()):
            self.assertTrue(self.app.open_path("chosen.WEX"))
        self.root.update()

    def select_group(self, index):
        self.ws.group_tabs.select(self.ws.group_tabs.tabs()[index])
        self.root.update()

    def test_explicit_import_and_group_switch_preserve_edits(self):
        self.assertIsNone(self.ws.proj)
        self.assertEqual(self.ws.group_tabs.tabs(), ())
        self.load_project()
        labels = [self.ws.group_tabs.tab(tab, "text")
                  for tab in self.ws.group_tabs.tabs()]
        self.assertEqual(labels, ["1.First", "2.Second", "3.Empty"])
        self.assertFalse(self.ws._has_unsaved_changes())

        self.ws.sheet.set_cell_data(0, 1, "Edited")
        self.select_group(1)
        self.assertEqual(self.ws.tab_rows[1][0]["Music_name"], "Edited")
        self.assertTrue(self.ws._has_unsaved_changes())
        self.select_group(0)
        self.assertEqual(self.ws.sheet.get_sheet_data()[0][1], "Edited")

    def test_empty_group_does_not_count_as_an_edit(self):
        self.load_project()
        self.select_group(2)
        self.assertFalse(self.ws._has_unsaved_changes())
        self.select_group(0)
        self.assertFalse(self.ws._has_unsaved_changes())

    def test_selected_group_has_distinct_tab_colors(self):
        self.load_project()
        style = ttk.Style()
        tab_style = self.ws.group_tabs.cget("style") + ".Tab"
        selected_bg = style.lookup(tab_style, "background", ("selected",))
        normal_bg = style.lookup(tab_style, "background")
        selected_fg = style.lookup(tab_style, "foreground", ("selected",))
        normal_fg = style.lookup(tab_style, "foreground")
        self.assertNotEqual(selected_bg, normal_bg)
        self.assertNotEqual(selected_fg, normal_fg)
        self.select_group(1)
        self.assertEqual(self.ws.current_tab, 2)

    def test_import_button_uses_the_selected_file(self):
        with patch("wex_editor.filedialog.askopenfilename", return_value="chosen.WEX"):
            with patch("wex_editor.WexProject", return_value=FakeProject()) as loader:
                self.ws.import_old_wex()
        loader.assert_called_once_with("chosen.WEX")
        self.assertEqual(self.ws.path, "chosen.WEX")
        with patch("wex_editor.filedialog.askopenfilename", return_value=""):
            with patch("wex_editor.WexProject") as loader:
                self.ws.import_old_wex()
                loader.assert_not_called()
        self.assertEqual(self.ws.path, "chosen.WEX")

    def test_cancelled_export_and_import_preserve_unsaved_changes(self):
        self.load_project()
        self.ws.sheet.set_cell_data(0, 1, "Edited")
        with patch("wex_editor.filedialog.asksaveasfilename", return_value=""):
            self.assertFalse(self.ws.export_dialog())
        self.assertTrue(self.ws._has_unsaved_changes())
        with patch("wex_editor.messagebox.askyesnocancel", return_value=None):
            with patch("wex_editor.WexProject") as loader:
                self.assertFalse(self.app.open_path("another.WEX"))
                loader.assert_not_called()
        self.assertEqual(self.ws.sheet.get_sheet_data()[0][1], "Edited")

    def test_export_requires_explicit_wex(self):
        with patch("wex_editor.messagebox.showerror") as error:
            self.assertFalse(self.ws.export_dialog())
        error.assert_called_once()
        self.assertIsNone(self.ws.proj)

    def import_exam(self, result, rows=None):
        if rows is None:
            rows = [
                dict(schedule_row(-1, "Exam one").data, Dir0="tone.wav"),
                dict(schedule_row(-2, "Exam two").data, Dir0="tone.wav"),
            ]
        with patch("wex_editor.filedialog.askopenfilename", return_value="exam.xlsx"):
            with patch.object(self.ws, "_load_invigilation_rows", return_value=rows):
                with patch("wex_editor.ScheduleImportDialog",
                           return_value=SimpleNamespace(result=result)):
                    with patch.object(self.ws, "_pref_auto_merge", return_value=False):
                        self.ws.import_schedule_excel()
        self.root.update()

    def test_import_requires_an_explicit_wex(self):
        with patch("wex_editor.messagebox.showinfo") as info:
            with patch("wex_editor.filedialog.askopenfilename") as file_dialog:
                self.ws.import_schedule_excel()
        info.assert_called_once()
        file_dialog.assert_not_called()

    def test_import_can_replace_another_group_and_preserves_current_edits(self):
        self.load_project()
        self.ws.sheet.set_cell_data(0, 1, "Unsaved first")
        self.import_exam(("replace", 2, "Second"))
        self.assertEqual(self.ws.current_tab, 2)
        self.assertEqual(self.ws.tab_rows[1][0]["Music_name"], "Unsaved first")
        self.assertEqual([r["Music_name"] for r in self.ws.tab_rows[2]],
                         ["Exam one", "Exam two"])
        self.assertEqual(self.ws.tab_rows[3], [])
        self.assertEqual(self.ws.tab_new_seq[2], -2)
        sql = self.ws._build_statements()
        self.assertTrue(any("DELETE FROM [MUSIC2]" in s for s in sql))
        self.assertTrue(any("DELETE FROM [MUSIC_T]" in s for s in sql))
        self.assertEqual(sum("INSERT INTO [MUSIC_T]" in s for s in sql), 2)
        self.assertTrue(any("[Up_music2]" in s for s in sql))
        self.assertFalse(any("[Up_page]" in s for s in sql))

    def test_import_can_create_named_group_and_updates_name_timestamp(self):
        self.load_project()
        self.import_exam(("new", 4, "New group"))
        self.assertEqual(self.ws.current_tab, 4)
        self.assertEqual(self.ws.tab_names[4], "New group")
        self.assertEqual([r["MUSIC"] for r in self.ws.tab_rows[4]], [-1, -2])
        self.assertEqual([r["Music_name"] for r in self.ws.tab_rows[2]], ["Second row"])
        sql = self.ws._build_statements()
        self.assertTrue(any("SELECT * INTO [MUSIC4]" in s for s in sql))
        self.assertTrue(any("[Up_music4]" in s for s in sql))
        self.assertTrue(any("[Up_music]" in s for s in sql))
        self.assertFalse(any("[Up_page]" in s for s in sql))

    def test_cancelled_import_does_not_change_the_sheet(self):
        self.load_project()
        self.ws.sheet.set_cell_data(0, 1, "Unsaved first")
        self.import_exam(None)
        self.assertEqual(self.ws.current_tab, 1)
        self.assertEqual(self.ws.sheet.get_sheet_data()[0][1], "Unsaved first")
        self.assertEqual(self.ws.tab_rows[2][0]["Music_name"], "Second row")

    def test_missing_audio_prevents_import(self):
        self.load_project()
        rows = [dict(schedule_row(-1, "Exam").data, Dir0="missing.wav")]
        with patch("wex_editor.messagebox.showerror") as error:
            self.import_exam(("replace", 2, "Second"), rows)
        error.assert_called_once()
        self.assertEqual(self.ws.tab_rows[2][0]["Music_name"], "Second row")

    def test_merge_keeps_different_explicit_dates(self):
        self.load_project()
        rows = [
            dict(schedule_row(-1, "10/28-1-Exam").data, Dir0="tone.wav",
                 Clock_year="2026", Clock_month="10", Clock_day="28"),
            dict(schedule_row(-2, "10/29-1-Exam").data, Dir0="tone.wav",
                 Clock_year="2026", Clock_month="10", Clock_day="29"),
        ]
        self.import_exam(("replace", 2, "Second"), rows)
        self.ws.merge_entries(auto=True)
        self.assertEqual(len(self.ws.tab_rows[2]), 2)

    def test_merge_still_combines_weekly_rows(self):
        self.load_project()
        rows = [
            dict(schedule_row(-1, "10/28-1-Exam").data, Dir0="tone.wav", Week1=True),
            dict(schedule_row(-2, "10/29-1-Exam").data, Dir0="tone.wav", Week2=True),
        ]
        self.import_exam(("replace", 2, "Second"), rows)
        self.ws.merge_entries(auto=True)
        self.assertEqual(len(self.ws.tab_rows[2]), 1)
        self.assertTrue(self.ws.tab_rows[2][0]["Week1"])
        self.assertTrue(self.ws.tab_rows[2][0]["Week2"])

    def test_ninth_group_requires_confirmation(self):
        def finish_dialog():
            dialog = next(c for c in self.root.winfo_children()
                          if isinstance(c, ScheduleImportDialog))
            dialog.mode.set("new")
            dialog._update_controls()
            dialog.new_name.set("New group")
            dialog.ok()
        self.root.after(50, finish_dialog)
        with patch("wex_editor.messagebox.askyesno", return_value=True) as confirm:
            dialog = ScheduleImportDialog(self.root, [(1, "First")], 1, 9, 50)
        confirm.assert_called_once()
        self.assertEqual(dialog.result, ("new", 9, "New group"))

    def test_declined_overwrite_does_not_select_a_destination(self):
        def finish_dialog():
            dialog = next(c for c in self.root.winfo_children()
                          if isinstance(c, ScheduleImportDialog))
            dialog.ok()
            dialog.cancel()
        self.root.after(50, finish_dialog)
        with patch("wex_editor.messagebox.askyesno", return_value=False) as confirm:
            dialog = ScheduleImportDialog(self.root, [(1, "First")], 1, 2, 50)
        confirm.assert_called_once()
        self.assertIsNone(dialog.result)

    def test_new_group_requires_a_unique_nonempty_name(self):
        def finish_dialog():
            dialog = next(c for c in self.root.winfo_children()
                          if isinstance(c, ScheduleImportDialog))
            dialog.mode.set("new")
            dialog._update_controls()
            dialog.ok()
            dialog.new_name.set("first")
            dialog.ok()
            dialog.cancel()
        self.root.after(50, finish_dialog)
        with patch("wex_editor.messagebox.showerror") as error:
            dialog = ScheduleImportDialog(self.root, [(1, "First")], 1, 2, 50)
        self.assertEqual(error.call_count, 2)
        self.assertIsNone(dialog.result)

    def test_many_groups_can_be_scrolled(self):
        self.ws.visible = list(range(1, 21))
        self.ws.tab_names = {i: f"Long schedule group {i}" for i in self.ws.visible}
        self.ws.current_tab = 1
        self.ws._refresh_group_tabs()
        self.root.update()
        self.ws.tab_canvas.xview_moveto(1)
        self.root.update()
        self.assertGreater(self.ws.tab_canvas.xview()[0], 0)
        self.assertEqual(len(self.ws.group_tabs.tabs()), 20)


if __name__ == "__main__":
    unittest.main()
