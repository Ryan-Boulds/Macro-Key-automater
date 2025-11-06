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
from ui_components import render_section, render_gap_chip, add_typed_dialog, step_label


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

        self.step_labels = []
        self.gap_chips = []
        self.row_gap_chips = []
        self.row_canvases = []  # (canvas, inner_frame)
        self.last_recorded_step = None
        self.active_row_index = 0
        self.active_section_index = None

        self.pending_update = False
        self.append_after_id = None
        self.pending_steps = []

        # === LAYOUT ===
        self.top_bar = TopBar(self.root, self)
        self.top_bar.frame.pack(side="top", fill="x")

        self._setup_canvas()

        if os.path.exists("temp_macro.json"):
            self.load_macro("temp_macro.json")
        if not self.recorder.rows:
            self.add_row()
        self.render_sections()

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.root.bind("<FocusIn>", lambda e: self._on_focus())
        self.root.bind("<Map>", lambda e: self._on_focus())

    def _setup_canvas(self):
        self.main_canvas = tk.Canvas(self.root)
        self.v_scrollbar = tk.Scrollbar(self.root, orient="vertical", command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=self.v_scrollbar.set)

        self.v_scrollbar.pack(side="right", fill="y")
        self.main_canvas.pack(side="top", fill="both", expand=True)

        self.h_scrollbar = tk.Scrollbar(self.root, orient="horizontal", command=self._h_scroll)
        self.h_scrollbar.pack(side="bottom", fill="x")

        self.main_frame = tk.Frame(self.main_canvas)
        self.main_canvas.create_window((0, 0), window=self.main_frame, anchor="nw", tags="inner")

        self.main_frame.bind("<Configure>", self._on_frame_configure)
        self.main_canvas.bind("<Configure>", self._on_canvas_configure)

    def _on_frame_configure(self, event=None):
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.main_canvas.itemconfig("inner", width=event.width)

    def _h_scroll(self, *args):
        for row_canvas, _ in self.row_canvases:
            row_canvas.xview(*args)

    def _row_xscroll(self, *args):
        if not self.row_canvases:
            return
        first = self.row_canvases[0][0]
        if first.xview() != (float(args[0]), float(args[1])):
            for canvas, _ in self.row_canvases:
                canvas.xview_moveto(args[0])
        self.h_scrollbar.set(*args)

    def _sync_h_scrollbars(self):
        if self.row_canvases:
            x0, x1 = self.row_canvases[0][0].xview()
            self.h_scrollbar.set(x0, x1)

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
        steps = self.pending_steps[:]
        self.pending_steps.clear()
        self.append_after_id = None

        for r, s, step, idx in steps:
            if r >= len(self.row_canvases) or s >= len(self.step_labels[r]):
                continue
            container = self.row_canvases[r][1]
            lbl = tk.Label(
                container,
                text=step_label(step),
                anchor="w", justify="left",
                bg="white", relief="ridge", bd=1, padx=4, pady=2
            )
            lbl.pack(side="top", fill="x", pady=1)
            self.step_labels[r][s].append(lbl)
            self._bind_step(lbl, (r, s), len(self.step_labels[r][s]) - 1)

        self.ui.refresh()
        self._sync_h_scrollbars()
        for canvas, _ in self.row_canvases:
            canvas.configure(scrollregion=canvas.bbox("all"))

    def _bind_step(self, lbl, si, sti):
        def on_click(e):
            ctrl_held = bool(e.state & 0x4)
            self.selection.toggle(si, sti, lbl, ctrl_held)
        lbl.bind("<Button-1>", on_click)

    def render_sections(self):
        if self.recorder.recording and self.active_section_index:
            return

        for widget in self.main_frame.winfo_children():
            widget.destroy()

        self.step_labels = []
        self.gap_chips = []
        self.row_gap_chips = []
        self.row_canvases = []

        rows = self.recorder.snapshot_rows()
        within = self.recorder.snapshot_within_rows()
        between = self.recorder.snapshot_between_rows()

        for r_idx, row in enumerate(rows):
            self.step_labels.append([])
            self.gap_chips.append([])

            row_canvas = tk.Canvas(self.main_frame, height=320, highlightthickness=0)
            row_frame = tk.Frame(row_canvas)
            row_canvas.create_window((0, 0), window=row_frame, anchor="nw")
            row_frame.bind("<Configure>", lambda e, c=row_canvas: c.configure(scrollregion=c.bbox("all")))

            row_canvas.configure(xscrollcommand=self._row_xscroll)
            row_canvas.pack(side="top", fill="x", padx=10, pady=8)
            self.row_canvases.append((row_canvas, row_frame))

            for s_idx, section in enumerate(row):
                self.step_labels[r_idx].append([])
                # Pass row_frame instead of app.sections_frame
                sec_frame = render_section(self, (r_idx, s_idx), section, parent=row_frame)
                sec_frame.pack(side="left", padx=8)

                if s_idx < len(row) - 1:
                    gap_frame = render_gap_chip(self, (r_idx, s_idx), within[r_idx][s_idx])
                    gap_frame.pack(in_=row_frame, side="left", padx=4)

            if r_idx < len(rows) - 1:
                gap_frame = render_gap_chip(self, (r_idx, -1), between[r_idx])
                gap_frame.pack(in_=self.main_frame, side="top", pady=6)
                self.row_gap_chips.append(gap_frame)

        self.ui.refresh()
        self._on_frame_configure()
        self._sync_h_scrollbars()

    # === Rest of methods unchanged ===
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
            self.root.after(100, self.render_sections)
        else:
            if self.active_section_index is None:
                messagebox.showerror("Error", "Select a section first.")
                return
            r, s = self.active_section_index
            self.step_labels[r][s] = []
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
            if ms < 0:
                raise ValueError
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
        try:
            os.remove("temp_macro.json")
        except:
            pass
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

    def _playback_highlight(self, section_idx, step_idx, active):
        pass

    def _on_closing(self):
        if os.path.exists("temp_macro.json"):
            try:
                os.remove("temp_macro.json")
            except:
                pass
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = MacroEditorApp(root)
    root.mainloop()