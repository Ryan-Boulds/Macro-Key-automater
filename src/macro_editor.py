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
        self.recorder.ui_callback = self._ui_callback
        self.recorder.playback_ui_callback = self._playback_highlight

        self.playback = PlaybackManager(self)
        self.selection = SelectionManager(self)
        self.section = SectionManager(self)
        self.ui = UIUpdater(self)

        self.step_labels = []          # step_labels[row][sec][step]
        self.step_menus = []
        self.gap_chips = []            # within-row gaps
        self.row_gap_chips = []        # between-row gaps
        self.last_recorded_step = None # (row, sec, step)
        self.active_row_index = 0
        self.active_section_index = None  # (row, sec)

        self.pending_update = False  # New: Flag from old code to defer renders

        # ---- UI setup first (so sections_frame exists) ----
        self.top_bar = TopBar(root, self)
        self._setup_canvas()          # creates self.sections_frame

        # ---- Now load temp file (safe) ----
        if os.path.exists("temp_macro.json"):
            self.load_macro("temp_macro.json")

        # ---- Guarantee at least one row ----
        if not self.recorder.rows:
            self.add_row()

        self.render_sections()        # **initial UI build**

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.root.bind("<FocusIn>", lambda e: self._on_focus())
        self.root.bind("<Map>", lambda e: self._on_focus())

    # New: Visibility check from old code
    def _is_visible(self):
        return self.root.winfo_viewable() and self.root.state() == "normal"

    # New: Handle deferred updates on focus/deiconify, from old code
    def _on_focus(self):
        self.ui.on_focus()
        if self.pending_update and self._is_visible():
            self.render_sections()
            self.pending_update = False

    # Updated: Thread-safe, conditional UI callback inspired by old code
    def _ui_callback(self):
        def scheduled_update():
            if self._is_visible():
                if self.recorder.recording:
                    # During recording: Incremental append to active section for speed
                    self._append_new_steps_to_ui()
                else:
                    self.render_sections()  # Full render only if not recording
            else:
                self.pending_update = True  # Defer if minimized
        self.root.after(0, scheduled_update)  # Main thread safe

    # New: Incremental UI append during recording (avoids full rebuild)
    def _append_new_steps_to_ui(self):
        if self.active_section_index is None:
            return
        r, s = self.active_section_index
        section = self.recorder.rows[r][s]
        current_steps = section["steps"]
        existing_labels = self.step_labels[r][s] if r < len(self.step_labels) and s < len(self.step_labels[r]) else []
        
        # Append only new labels since last render
        for idx in range(len(existing_labels), len(current_steps)):
            step = current_steps[idx]
            lbl = tk.Text(self.sections_frame, width=STEP_WIDTH, height=STEP_HEIGHT, wrap="word", bd=1, relief="ridge")
            lbl.insert("1.0", step_label(step))
            lbl.config(state="disabled")
            lbl.pack(fill="x", pady=2)
            self.step_labels[r][s].append(lbl)
            # Bind clicks, etc. (simplified; add full bindings as in render_section)
            lbl.bind("<Button-1>", lambda e, si=(r,s), sti=idx: self.selection.toggle(si, sti, lbl, e.state & 0x4))  # Ctrl check
        self.ui.refresh()  # Update highlights without full rebuild

    # ------------------------------------------------------------------ #
    # Canvas / scrolling (unchanged)
    # ------------------------------------------------------------------ #
    def _setup_canvas(self):
        outer = tk.Frame(self.root)
        outer.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(outer)
        v = tk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        h = tk.Scrollbar(outer, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=v.set, xscrollcommand=h.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        v.grid(row=0, column=1, sticky="ns")
        h.grid(row=1, column=0, sticky="ew")
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)

        self.sections_frame = tk.Frame(self.canvas)
        self.canvas_window_id = self.canvas.create_window((0, 0),
                                                          window=self.sections_frame,
                                                          anchor="nw")

        self.sections_frame.bind("<Configure>",
                                 lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>",
                         lambda e: self.canvas.itemconfig(self.canvas_window_id, width=e.width))

        self._bind_mousewheel()

    def _bind_mousewheel(self):
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.canvas.bind_all(seq, self._on_mousewheel)

    def _on_mousewheel(self, event):
        if event.num == 4 or event.delta > 0:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.canvas.yview_scroll(1, "units")

    # ------------------------------------------------------------------ #
    # UI callbacks
    # ------------------------------------------------------------------ #
    def _ui_callback(self):
        self.render_sections()

    def _playback_highlight(self, sec_idx, step_idx, active):
        self.ui.highlight_playback(sec_idx, step_idx, active)

    def _on_closing(self):
        self.save_temp_macro()
        self.root.destroy()

    def save_temp_macro(self):
        try:
            with open("temp_macro.json", "w") as f:
                json.dump({
                    "rows": self.recorder.rows,
                    "delays_within_rows": self.recorder.delays_within_rows,
                    "delays_between_rows": self.recorder.delays_between_rows
                }, f)
        except:
            pass

    # ------------------------------------------------------------------ #
    # Rendering – rebuilds the whole UI every time the model changes
    # ------------------------------------------------------------------ #
    def render_sections(self):
        # destroy old widgets
        for menu in self.step_menus:
            try: menu.destroy()
            except: pass
        self.step_menus = []
        self.step_labels = [[[] for _ in row] for row in self.recorder.rows]
        self.gap_chips = [[] for _ in self.recorder.rows]
        self.row_gap_chips = []
        for w in self.sections_frame.winfo_children():
            w.destroy()

        snapshot = self.recorder.snapshot_rows()
        within = self.recorder.snapshot_within_rows()
        between = self.recorder.snapshot_between_rows()

        for row_idx, row in enumerate(snapshot):
            row_frame = tk.Frame(self.sections_frame)
            row_frame.pack(fill="x", pady=8, padx=8)

            content = tk.Frame(row_frame)
            content.pack(fill="x")

            for sec_idx, sec in enumerate(row):
                sec_frame = render_section(self, (row_idx, sec_idx), sec)
                sec_frame.pack(side="left", padx=8)

                if sec_idx < len(row) - 1:
                    gap = render_gap_chip(self, (row_idx, sec_idx),
                                          within[row_idx][sec_idx])
                    gap.pack(side="left", padx=(0, 8))

            # between-row delay chip
            if row_idx < len(snapshot) - 1:
                delay_frame = tk.Frame(self.sections_frame)
                delay_frame.pack(fill="x", pady=4, padx=8)
                tk.Label(delay_frame, text="Delay Between Rows (ms):").pack(side="left")
                var = tk.StringVar(value=str(between[row_idx]))
                tk.Entry(delay_frame, textvariable=var, width=6).pack(side="left", padx=4)
                tk.Button(delay_frame, text="Set",
                          command=lambda ri=row_idx, v=var: self.set_between_row_delay(ri, v.get())
                         ).pack(side="left")
                self.row_gap_chips.append(delay_frame)

    def set_between_row_delay(self, row_idx, value):
        try:
            ms = int(float(value))
            if ms < 0: raise ValueError
            self.recorder.set_between_row_delay(row_idx, ms)
            self.render_sections()
        except ValueError:
            messagebox.showerror("Error", "Enter a valid delay (ms).")

    # ------------------------------------------------------------------ #
    # Public actions (called from TopBar / UI)
    # ------------------------------------------------------------------ #
    def add_row(self):
        self.recorder.add_row()
        self.active_row_index = len(self.recorder.rows) - 1
        self.render_sections()

    def add_section(self):
        """Add Column button → new section in the active row."""
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
            if self.active_section_index:
                r, s = self.active_section_index
                steps = self.recorder.rows[r][s]["steps"]
                self.last_recorded_step = (*self.active_section_index,
                                          len(steps) - 1) if steps else None
        else:
            if self.active_section_index is None:
                messagebox.showerror("Error", "Select a section first.")
                return
            self.last_recorded_step = None
            self.recorder.start_recording(self.active_section_index)
            self.top_bar.update_record_button(True)
            if self.top_bar.get_auto_minimize():
                self.root.iconify()
        self.render_sections()

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
        self.render_sections()

    def move_step_up(self, sec_idx, step_idx):
        self.recorder.move_step_up(sec_idx, step_idx)
        self.render_sections()

    def move_step_down(self, sec_idx, step_idx):
        self.recorder.move_step_down(sec_idx, step_idx)
        self.render_sections()


if __name__ == "__main__":
    root = tk.Tk()
    app = MacroEditorApp(root)
    root.mainloop()