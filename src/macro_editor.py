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
        self.selected_steps = self.selection.selected_steps
        self.section = SectionManager(self)
        self.ui = UIUpdater(self)

        self.step_labels = []
        self.step_menus = []
        self.gap_chips = []
        self.last_recorded_step = None
        self.active_section_index = 0

        # Load temp
        if os.path.exists("temp_macro.json"):
            self.load_macro("temp_macro.json")

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.root.bind("<FocusIn>", lambda e: self.ui.on_focus())
        self.root.bind("<Map>", lambda e: self.ui.on_focus())

        # UI
        self.top_bar = TopBar(root, self)
        self._setup_canvas()
        if not self.recorder.sections:
            self.section.add("Section 1")
        self.render_sections()

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
        self.canvas_window_id = self.canvas.create_window((0, 0), window=self.sections_frame, anchor="nw")

        self.sections_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self.canvas_window_id, width=e.width))

        self._bind_mousewheel()

    def _bind_mousewheel(self):
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.canvas.bind_all(seq, self._on_mousewheel)

    def _on_mousewheel(self, event):
        if event.num == 4 or event.delta > 0:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.canvas.yview_scroll(1, "units")

    def _ui_callback(self):
        self.ui.refresh()

    def _playback_highlight(self, sec_idx, step_idx, active):
        self.ui.highlight_playback(sec_idx, step_idx, active)

    def _is_visible(self):
        return self.root.state() != 'iconic' and self.root.focus_get() is not None

    def _on_closing(self):
        self.save_temp_macro()
        self.root.destroy()

    def save_temp_macro(self):
        try:
            with open("temp_macro.json", "w") as f:
                json.dump({"sections": self.recorder.sections, "delays_between": self.recorder.delays_between}, f)
        except: pass

    def render_sections(self):
        # Clear old
        for menu in self.step_menus:
            try: menu.destroy()
            except: pass
        self.step_menus = []
        self.step_labels = [[] for _ in self.recorder.sections]
        self.gap_chips = []
        for w in self.sections_frame.winfo_children():
            w.destroy()

        sections = self.recorder.snapshot_sections()
        gaps = self.recorder.snapshot_between_delays()

        col = 0
        for idx, sec in enumerate(sections):
            frame = render_section(self, idx, sec)
            frame.grid(row=0, column=col, padx=8, pady=8, sticky="n")
            col += 1
            if idx < len(sections) - 1:
                gap_frame = render_gap_chip(self, idx, gaps[idx])
                gap_frame.grid(row=0, column=col, padx=(0,0), pady=8, sticky="ns")
                col += 1

    # --- Delegate methods ---
    def add_section(self): self.section.add()
    def add_string(self):
        if self.active_section_index is None: messagebox.showerror("Error", "Select a section first."); return
        add_typed_dialog(self, self.active_section_index)
    def toggle_recording(self):
        if self.recorder.recording:
            self.recorder.stop_recording()
            self.top_bar.update_record_button(False)
            if self.active_section_index is not None:
                steps = self.recorder.snapshot_sections()[self.active_section_index]["steps"]
                self.last_recorded_step = (self.active_section_index, len(steps)-1) if steps else None
        else:
            if self.active_section_index is None: messagebox.showerror("Error", "Select a section first."); return
            self.last_recorded_step = None
            self.recorder.start_recording(self.active_section_index)
            self.top_bar.update_record_button(True)
            if self.top_bar.get_auto_minimize():
                self.root.iconify()
        self.ui.refresh()

    def play_macro(self): self.playback.start_playback()
    def add_quick_delay(self):
        if self.active_section_index is None: messagebox.showerror("Error", "Select a section first."); return
        try:
            ms = int(float(self.top_bar.quick_delay_var.get()))
            if ms < 0: raise ValueError
        except: messagebox.showerror("Error", "Enter a valid delay (ms)."); return
        self.recorder.add_delay_step(self.active_section_index, ms)

    def save_macro(self, file=None):
        if not file: file = filedialog.asksaveasfilename(defaultextension=".json")
        if file: self.recorder.save_macro(file); messagebox.showinfo("Save", "Saved.")

    def load_macro(self, file=None):
        if not file: file = filedialog.askopenfilename()
        if file:
            self.recorder.load_macro(file)
            self.last_recorded_step = None
            self.selection.clear()
            self.ui.refresh()

    def clear_all(self):
        self.recorder.clear_all()
        self.last_recorded_step = None
        self.selection.clear()
        try: os.remove("temp_macro.json")
        except: pass
        self.ui.refresh()

    def select_section(self, idx): self.section.select(idx)
    def delete_section(self, idx): self.section.delete(idx)
    def move_section_left(self, idx): self.section.move_left(idx)
    def move_section_right(self, idx): self.section.move_right(idx)
    def delete_step(self, si, sti):
        self.selection.clear()
        self.recorder.delete_step(si, sti)
        if (si, sti) == self.last_recorded_step:
            self.last_recorded_step = None
        self.ui.refresh()
    def move_step_up(self, si, sti):
        self.selection.clear()
        self.recorder.move_step_up(si, sti)
        self.ui.refresh()
    def move_step_down(self, si, sti):
        self.selection.clear()
        self.recorder.move_step_down(si, sti)
        self.ui.refresh()

if __name__ == "__main__":
    root = tk.Tk()
    app = MacroEditorApp(root)
    root.mainloop()