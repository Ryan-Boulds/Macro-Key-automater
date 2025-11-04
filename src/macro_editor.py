# src/macro_editor.py
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json

from macro_recorder import MacroRecorderCore
from top_bar import TopBar
from playback_manager import PlaybackManager
from selection_manager import SelectionManager
from section_manager import SectionManager
from ui_updater import UIUpdater
from ui_components import render_section, render_gap_chip, add_typed_dialog


class MacroEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Macro Recorder")
        self.root.geometry("1200x700")

        self.recorder = MacroRecorderCore()
        self.recorder.root = self.root
        self.recorder.ui_callback = self._ui_callback
        self.recorder.playback_ui_callback = self._playback_highlight

        self.playback = PlaybackManager(self)
        self.selection = SelectionManager(self)
        self.section = SectionManager(self)
        self.ui = UIUpdater(self)

        self.step_labels = []          # step_labels[row][sec] = list of Labels
        self.step_menus = []
        self.gap_chips = []            # within-row gaps
        self.row_gap_chips = []        # between-row gaps
        self.last_recorded_step = None # (row, sec, step)
        self.active_row_index = 0
        self.active_section_index = None  # (row, sec)

        self.pending_update = False
        self.append_after_id = None    # For batching appends
        self.pending_steps = []        # Buffer for batched append

        self._setup_canvas()
        if os.path.exists("temp_macro.json"):
            self.load_macro("temp_macro.json")
        if not self.recorder.rows:
            self.add_row()
        self.render_sections()

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.root.bind("<FocusIn>", lambda e: self._on_focus())
        self.root.bind("<Map>", lambda e: self._on_focus())

    def _is_visible(self):
        return self.root.winfo_viewable() and self.root.state() == "normal"

    def _on_focus(self):
        self.ui.on_focus()
        if self.pending_update and self._is_visible():
            self.render_sections()
            self.pending_update = False

    def _ui_callback(self):
        if not self._is_visible():
            self.pending_update = True
            return
        if self.recorder.recording:
            self.root.after(0, self._flush_pending_steps)
        else:
            self.root.after(0, self.render_sections)

    def _queue_step_for_append(self, r, s, step, idx):
        self.pending_steps.append((r, s, step, idx))
        if not self.append_after_id:
            self.append_after_id = self.root.after(50, self._flush_pending_steps)

    def _flush_pending_steps(self):
        if not self.pending_steps:
            self.append_after_id = None
            return
        steps_to_add = self.pending_steps
        self.pending_steps = []
        self.append_after_id = None

        for r, s, step, idx in steps_to_add:
            if (r >= len(self.step_labels) or s >= len(self.step_labels[r]) or
                idx < len(self.step_labels[r][s])):
                continue
            lbl = tk.Label(
                self.sections_frame,
                text=step_label(step),
                anchor="w",
                justify="left",
                bg="white",
                relief="ridge",
                bd=1,
                padx=4,
                pady=2
            )
            lbl.pack(fill="x", pady=1)
            self.step_labels[r][s].append(lbl)
            self._bind_step(lbl, (r, s), idx)
        self.ui.refresh()

    def _bind_step(self, lbl, si, sti):
        def on_click(e, shift=tk.SHIFT, ctrl=tk.CTRL):
            ctrl_held = bool(e.state & ctrl)
            self.selection.toggle(si, sti, lbl, ctrl_held)
        lbl.bind("<Button-1>", on_click)

    def _setup_canvas(self):
        canvas = tk.Canvas(self.root)
        scrollbar = tk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self.sections_frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=self.sections_frame, anchor="nw")
        self.sections_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        self.top_bar = TopBar(self.root, self)
        self.top_bar.frame.pack(fill="x")

    def render_sections(self):
        if self.recorder.recording and self.active_section_index:
            return  # Skip full render during recording
        for widget in self.sections_frame.winfo_children():
            widget.destroy()
        self.step_labels = []
        self.gap_chips = []
        self.row_gap_chips = []

        rows = self.recorder.snapshot_rows()
        within = self.recorder.snapshot_within_rows()
        between = self.recorder.snapshot_between_rows()

        for r_idx, row in enumerate(rows):
            self.step_labels.append([])
            self.gap_chips.append([])
            row_frame = tk.Frame(self.sections_frame)
            row_frame.pack(fill="x", pady=8)

            for s_idx, section in enumerate(row):
                self.step_labels[r_idx].append([])
                sec_frame = render_section(self, (r_idx, s_idx), section)
                sec_frame.pack(side="left", padx=8)
                if s_idx < len(row) - 1:
                    gap_frame = render_gap_chip(self, (r_idx, s_idx), within[r_idx][s_idx])
                    gap_frame.pack(side="left", padx=4)

            if r_idx < len(rows) - 1:
                gap_frame = render_gap_chip(self, (r_idx, -1), between[r_idx])
                gap_frame.pack(in_=row_frame, side="bottom", pady=4)
                self.row_gap_chips.append(gap_frame)

        self.ui.refresh()

    def add_row(self):
        self.recorder.add_row()
        self.active_row_index = len(self.recorder.rows) - 1
        self.render_sections()

    def add_section(self):
        if not self.recorder.rows:
            self.add_row()
        name = f"Section {len(self.recorder.rows[self.active_row_index]) + 1}"
        idx = self.recorder.add_section(self.active_row_index, name)
        self.active_section_index = idx
        self.render_sections()

    def add_string(self):
        if self.active_section_index is None:
            messagebox.showerror("Error", "Select a section first.")
            return
        add_typed_dialog(self, self.active_section_index)

    def toggle_recording(self):
        if self.recorder.recording:
            self.recorder.stop_recording()
            self.top_bar.update_record_button(False)
            self.last_recorded_step = None
            self.root.after(100, self.render_sections)  # Rebuild once after stop
        else:
            if self.active_section_index is None:
                messagebox.showerror("Error", "Select a section first.")
                return
            r, s = self.active_section_index
            self.step_labels[r][s] = []  # Clear for incremental
            self.recorder.start_recording(self.active_section_index)
            self.top_bar.update_record_button(True)
            if self.top_bar.get_auto_minimize():
                self.root.iconify()

    def play_macro(self):
        self.playback.start_playback()

    def add_quick_delay(self):
        if self.active_section_index is None:
            messagebox.showerror("Error", "Select a section first.")
            return
        try:
            ms = int(float(self.top_bar.quick_delay_var.get()))
            if ms < 0: raise ValueError
        except Exception:
            messagebox.showerror("Error", "Enter a valid delay (ms).")
            return
        self.recorder.add_delay_step(self.active_section_index, ms)
        if not self.recorder.recording:
            self.render_sections()

    def save_macro(self, file=None):
        if not file:
            file = filedialog.asksaveasfilename(defaultextension=".json")
        if file:
            self.recorder.save_macro(file)
            messagebox.showinfo("Save", "Saved.")

    def load_macro(self, file=None):
        if not file:
            file = filedialog.askopenfilename()
        if file:
            self.recorder.load_macro(file)
            self.last_recorded_step = None
            self.selection.clear()
            self.render_sections()

    def clear_all(self):
        self.recorder.clear_all()
        self.last_recorded_step = None
        self.selection.clear()
        try: os.remove("temp_macro.json")
        except: pass
        self.add_row()
        self.render_sections()

    def select_section(self, idx):
        self.active_row_index, self.active_section_index = idx[0], idx
        self.render_sections()

    def delete_section(self, idx):
        self.recorder.delete_section(idx)
        if self.active_section_index == idx:
            self.active_section_index = None
        self.render_sections()

    def move_section_left(self, idx):
        self.recorder.move_section_left(idx)
        self.render_sections()

    def move_section_right(self, idx):
        self.recorder.move_section_right(idx)
        self.render_sections()

    def delete_step(self, sec_idx, step_idx):
        self.selection.clear()
        self.recorder.delete_step(sec_idx, step_idx)
        if (self.last_recorded_step and
                self.last_recorded_step[:2] == sec_idx and
                self.last_recorded_step[2] == step_idx):
            self.last_recorded_step = None
        if not self.recorder.recording:
            self.render_sections()

    def move_step_up(self, sec_idx, step_idx):
        self.recorder.move_step_up(sec_idx, step_idx)
        if not self.recorder.recording:
            self.render_sections()

    def move_step_down(self, sec_idx, step_idx):
        self.recorder.move_step_down(sec_idx, step_idx)
        if not self.recorder.recording:
            self.render_sections()

    def _playback_highlight(self, section_idx, step_idx, active):
        pass  # Handled by UIUpdater

    def _on_closing(self):
        if os.path.exists("temp_macro.json"):
            try: os.remove("temp_macro.json")
            except: pass
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = MacroEditorApp(root)
    root.mainloop()