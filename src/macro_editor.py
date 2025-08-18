import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import threading
from macro_recorder import MacroRecorderCore
from pynput import keyboard

STEP_WIDTH = 18
STEP_HEIGHT = 2


class MacroEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Macro Recorder with Columns & Between Delays")
        self.root.geometry("1200x700")

        self.recorder = MacroRecorderCore()
        self.recorder.ui_callback = lambda: self.root.after(0, self.render_sections)
        self.recorder.playback_ui_callback = self._playback_highlight

        self.stop_event = None
        self.interrupt_listener = None

        self.step_labels = []
        self.step_menus = []
        self.gap_chips = []
        self.selected_steps = {}  # (section_idx, step_idx) -> label
        self.last_clicked = None  # Last clicked step for single-step movement

        # ===== Top controls (stay pinned) =====
        top = tk.Frame(root)
        top.pack(side="top", fill="x", pady=6)

        tk.Button(top, text="Add Column", command=self.add_section).pack(side="left", padx=4)
        self.record_button = tk.Button(top, text="Start Recording", command=self.toggle_recording)
        self.record_button.pack(side="left", padx=4)
        tk.Button(top, text="Play Macro", command=self.play_macro).pack(side="left", padx=4)
        tk.Button(top, text="Save", command=self.save_macro).pack(side="left", padx=4)
        tk.Button(top, text="Load", command=self.load_macro).pack(side="left", padx=4)
        tk.Button(top, text="Clear All", command=self.clear_all).pack(side="left", padx=4)

        self.quick_delay_var = tk.StringVar(value="250")
        tk.Label(top, text="Step Delay ms:").pack(side="left", padx=(16, 4))
        tk.Entry(top, textvariable=self.quick_delay_var, width=6).pack(side="left")
        tk.Button(top, text="Add Step Delay to Selected", command=self.add_quick_delay).pack(side="left", padx=4)

        self.auto_minimize_var = tk.BooleanVar(value=False)
        tk.Checkbutton(top, text="Auto-minimize when recording", variable=self.auto_minimize_var).pack(side="left", padx=8)

        # ===== Scrollable area (both directions) =====
        outer = tk.Frame(root)
        outer.pack(side="top", fill="both", expand=True)

        self.canvas = tk.Canvas(outer)
        self.vscroll = tk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        self.hscroll = tk.Scrollbar(outer, orient="horizontal", command=self.canvas.xview)

        self.canvas.configure(yscrollcommand=self.vscroll.set, xscrollcommand=self.hscroll.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vscroll.grid(row=0, column=1, sticky="ns")
        self.hscroll.grid(row=1, column=0, sticky="ew")

        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)

        self.sections_frame = tk.Frame(self.canvas)
        self.canvas_window_id = self.canvas.create_window((0, 0), window=self.sections_frame, anchor="nw")

        self.sections_frame.bind("<Configure>", self._on_sections_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Up>", self._on_arrow_key)
        self.canvas.bind("<Down>", self._on_arrow_key)

        self._bind_mousewheel(self.canvas)

        if not self.recorder.sections:
            self.recorder.add_section("Section 1")

        self.active_section_index = 0
        self.render_sections()

    def _on_sections_configure(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window_id, width=max(event.width, self.sections_frame.winfo_reqwidth()))

    def _bind_mousewheel(self, widget):
        widget.bind_all("<MouseWheel>", self._on_mousewheel)
        widget.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)
        widget.bind_all("<Button-4>", lambda e: self.canvas.yview_scroll(-3, "units"))
        widget.bind_all("<Button-5>", lambda e: self.canvas.yview_scroll(+3, "units"))

    def _on_mousewheel(self, event):
        delta = -1 * int(event.delta / 120) if event.delta else 0
        self.canvas.yview_scroll(delta, "units")

    def _on_shift_mousewheel(self, event):
        delta = -1 * int(event.delta / 120) if event.delta else 0
        self.canvas.xview_scroll(delta, "units")

    def _scroll_to_widget(self, widget):
        x = 0
        y = 0
        w = widget
        while w != self.sections_frame:
            x += w.winfo_x()
            y += w.winfo_y()
            w = w.master
        x += widget.winfo_width() / 2
        y += widget.winfo_height() / 2

        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        frame_width = self.sections_frame.winfo_width()
        frame_height = self.sections_frame.winfo_height()

        frac_x = max(0, min(1, (x - canvas_width / 2) / frame_width))
        frac_y = max(0, min(1, (y - canvas_height / 2) / frame_height))

        self.canvas.xview_moveto(frac_x)
        self.canvas.yview_moveto(frac_y)

    def render_sections(self):
        # Clean up existing menus
        for menu in self.step_menus:
            try:
                menu.destroy()
            except:
                pass
        self.step_menus = []
        self.step_labels = []
        self.gap_chips = []
        self.selected_steps = {}
        self.last_clicked = None

        for w in self.sections_frame.winfo_children():
            w.destroy()

        sections = self.recorder.snapshot_sections()
        gaps = self.recorder.snapshot_between_delays()

        self.step_labels = [[] for _ in sections]

        col = 0
        for idx, section in enumerate(sections):
            sec_frame = self._render_one_section(idx, section)
            sec_frame.grid(row=0, column=col, padx=8, pady=8, sticky="n")
            col += 1

            if idx < len(sections) - 1:
                gap_index = idx
                gap_frame = self._render_gap_chip(gap_index, gaps[gap_index] if gap_index < len(gaps) else 0)
                gap_frame.grid(row=0, column=col, padx=(0, 0), pady=8, sticky="ns")
                col += 1

        self._on_sections_configure()

    def _render_gap_chip(self, gap_index, value_ms):
        frame = tk.Frame(self.sections_frame)
        chip = tk.Frame(frame, bd=1, relief="ridge", bg="white")
        chip.pack(fill="y", expand=True, padx=2, pady=2)
        self.gap_chips.append(chip)

        tk.Label(chip, text="Between", font=("TkDefaultFont", 8)).pack(padx=6, pady=(6, 0))
        var = tk.StringVar(value=str(value_ms))
        entry = tk.Entry(chip, textvariable=var, width=6, justify="center")
        entry.pack(padx=6, pady=4)

        def apply():
            try:
                ms = int(float(var.get()))
            except ValueError:
                messagebox.showerror("Error", "Enter a valid delay (ms).")
                return
            self.recorder.set_between_delay(gap_index, ms)

        btn = tk.Button(chip, text="Set ms", command=apply)
        btn.pack(padx=6, pady=(0, 6))

        return frame

    def _render_one_section(self, idx, section):
        is_active = (idx == self.active_section_index)
        border_color = "#0078D7" if is_active else "#cccccc"

        frame = tk.Frame(self.sections_frame, bd=2, relief="groove", highlightthickness=2)
        frame.configure(highlightbackground=border_color, highlightcolor=border_color)

        header = tk.Frame(frame)
        header.pack(fill="x", padx=6, pady=6)

        name_var = tk.StringVar(value=section["name"])
        name_entry = tk.Entry(header, textvariable=name_var, justify="center")
        name_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        name_entry.bind("<Return>", lambda _e, i=idx, v=name_var: self.recorder.rename_section(i, v.get()))
        name_entry.bind("<FocusOut>", lambda _e, i=idx, v=name_var: self.recorder.rename_section(i, v.get()))

        tk.Button(header, text="←", width=3, command=lambda i=idx: self.move_section_left(i)).pack(side="left", padx=2)
        tk.Button(header, text="→", width=3, command=lambda i=idx: self.move_section_right(i)).pack(side="left", padx=2)
        record_btn = tk.Button(header, text="Record Here", command=lambda i=idx: self.select_section(i))
        if is_active:
            record_btn.config(bg="red")
        record_btn.pack(side="left", padx=6)
        tk.Button(header, text="Delete", command=lambda i=idx: self.delete_section(i)).pack(side="left", padx=6)

        steps_wrap = tk.Frame(frame)
        steps_wrap.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        for s_idx, step in enumerate(section["steps"]):
            self._render_step(steps_wrap, idx, s_idx, step)

        return frame

    def _render_step(self, parent, section_idx, step_idx, step):
        row = tk.Frame(parent)
        row.pack(fill="x", pady=2)

        text = self._step_label(step)
        lbl = tk.Label(row, text=text, bd=1, relief="solid", width=STEP_WIDTH, height=STEP_HEIGHT, anchor="center", bg="white")
        lbl.pack(side="left")
        self.step_labels[section_idx].append(lbl)

        def toggle_selection(event, si=section_idx, sti=step_idx):
            key = (si, sti)
            if event.state & 0x4:  # Control key held
                if key in self.selected_steps:
                    self.selected_steps[key].config(bg="white")
                    del self.selected_steps[key]
                else:
                    self.selected_steps[key] = lbl
                    lbl.config(bg="#D3D3D3")
            else:
                self.clear_selection()
                self.selected_steps[key] = lbl
                lbl.config(bg="#D3D3D3")
            self.last_clicked = key

        lbl.bind("<Button-1>", lambda e, si=section_idx, sti=step_idx: toggle_selection(e, si, sti))
        lbl.bind("<Control-Button-1>", lambda e, si=section_idx, sti=step_idx: toggle_selection(e, si, sti))

        menu = tk.Menu(self.root, tearoff=0)
        self.step_menus.append(menu)
        menu.add_command(label="Delete", command=lambda si=section_idx, sti=step_idx: self.delete_step(si, sti))
        if step.get("type") == "delay":
            menu.add_command(label="Edit Delay…", command=lambda si=section_idx, sti=step_idx: self.edit_delay(si, sti))
        lbl.bind("<Button-3>", lambda e, m=menu: m.post(e.x_root, e.y_root))

        tk.Button(row, text="X", width=2, command=lambda si=section_idx, sti=step_idx: self.delete_step(si, sti)).pack(side="left", padx=2)

        ctrl = tk.Frame(row)
        ctrl.pack(side="left", padx=4)
        tk.Button(ctrl, text="↑", width=2, command=lambda si=section_idx, sti=step_idx: self.move_step_up(si, sti)).pack(side="top")
        tk.Button(ctrl, text="↓", width=2, command=lambda si=section_idx, sti=step_idx: self.move_step_down(si, sti)).pack(side="top")

    def _step_label(self, step):
        t = step.get("type")
        if t == "delay":
            return f"Delay {step['delay']} {step.get('unit','ms')}"
        if t == "press":
            return f"{step['key']} (pressed)"
        if t == "release":
            return f"{step['key']} (released)"
        if t == "mouse_press":
            return f"Mouse {step['button']} press @ ({step['x']}, {step['y']})"
        if t == "mouse_release":
            return f"Mouse {step['button']} release @ ({step['x']}, {step['y']})"
        return "Unknown"

    def _playback_highlight(self, sec_idx, step_idx, active):
        def do_highlight():
            bg_color = "#ADD8E6" if active else "white"
            widget = None
            if step_idx >= 0:
                if 0 <= sec_idx < len(self.step_labels) and 0 <= step_idx < len(self.step_labels[sec_idx]):
                    lbl = self.step_labels[sec_idx][step_idx]
                    if (sec_idx, step_idx) not in self.selected_steps:
                        lbl.config(bg=bg_color)
                    widget = lbl
            else:
                gap_idx = sec_idx
                if 0 <= gap_idx < len(self.gap_chips):
                    chip = self.gap_chips[gap_idx]
                    chip.config(bg=bg_color)
                    widget = chip
            if active and widget:
                self._scroll_to_widget(widget)
        self.root.after(0, do_highlight)

    def clear_selection(self):
        for (si, sti), lbl in self.selected_steps.items():
            lbl.config(bg="white")
        self.selected_steps.clear()

    def _on_arrow_key(self, event):
        if not self.selected_steps:
            if self.last_clicked:
                si, sti = self.last_clicked
                if event.keysym == "Up":
                    self.move_step_up(si, sti)
                elif event.keysym == "Down":
                    self.move_step_down(si, sti)
            return
        self.move_selected_steps(event.keysym)

    def move_selected_steps(self, direction):
        if not self.selected_steps:
            return
        # Group selections by section
        sections = {}
        for (si, sti) in self.selected_steps:
            if si not in sections:
                sections[si] = []
            sections[si].append(sti)
        
        for si in sections:
            indices = sorted(sections[si])
            if direction == "Up" and indices[0] > 0:
                self.recorder.move_steps_up(si, indices)
            elif direction == "Down" and indices[-1] < len(self.recorder.snapshot_sections()[si]["steps"]) - 1:
                self.recorder.move_steps_down(si, indices)
        
        self.render_sections()

    def add_section(self):
        idx = self.recorder.add_section(f"Section {len(self.recorder.sections) + 0}")
        self.active_section_index = idx
        self.render_sections()

    def delete_section(self, idx):
        self.recorder.delete_section(idx)
        if self.active_section_index is not None:
            if self.active_section_index >= len(self.recorder.snapshot_sections()):
                self.active_section_index = max(0, len(self.recorder.snapshot_sections()) - 1)
        self.render_sections()

    def select_section(self, idx):
        self.active_section_index = idx
        self.recorder.active_section_index = idx
        self.render_sections()

    def move_section_left(self, idx):
        self.recorder.move_section_left(idx)

    def move_section_right(self, idx):
        self.recorder.move_section_right(idx)

    def delete_step(self, section_idx, step_idx):
        self.recorder.delete_step(section_idx, step_idx)

    def move_step_up(self, section_idx, step_idx):
        self.recorder.move_step_up(section_idx, step_idx)

    def move_step_down(self, section_idx, step_idx):
        self.recorder.move_step_down(section_idx, step_idx)

    def edit_delay(self, section_idx, step_idx):
        current = self.recorder.snapshot_sections()[section_idx]["steps"][step_idx]
        value = current.get("delay", 0)
        try:
            new_val = simpledialog.askinteger("Edit Delay", "Delay (ms):", initialvalue=int(value), minvalue=0)
            if new_val is not None:
                self.recorder.edit_delay(section_idx, step_idx, int(new_val))
        except Exception:
            pass

    def add_quick_delay(self):
        if self.active_section_index is None:
            messagebox.showerror("Error", "Select a section first.")
            return
        try:
            ms = int(float(self.quick_delay_var.get()))
        except ValueError:
            messagebox.showerror("Error", "Enter a valid delay (ms).")
            return
        self.recorder.add_delay_step(self.active_section_index, ms)

    def toggle_recording(self):
        if self.recorder.recording:
            self.recorder.stop_recording()
            self.record_button.config(text="Start Recording", bg="SystemButtonFace")
        else:
            if self.active_section_index is None or self.active_section_index >= len(self.recorder.sections):
                messagebox.showerror("Error", "Select a section first.")
                return
            new_name = simpledialog.askstring("Rename Section", "Enter new section name:", parent=self.root)
            if new_name and new_name.strip():
                self.recorder.rename_section(self.active_section_index, new_name.strip())
            self.recorder.start_recording(self.active_section_index)
            self.record_button.config(text="Stop Recording", bg="red")
            if self.auto_minimize_var.get():
                self.root.iconify()

    def play_macro(self):
        self.stop_event = threading.Event()
        self.pressed = set()

        def on_press_key(k):
            self.pressed.add(k)
            if {keyboard.Key.ctrl, keyboard.Key.alt, keyboard.Key.enter}.issubset(self.pressed):
                self.stop_event.set()

        def on_release_key(k):
            if k in self.pressed:
                self.pressed.remove(k)

        self.interrupt_listener = keyboard.Listener(on_press=on_press_key, on_release=on_release_key)
        self.interrupt_listener.start()
        threading.Thread(target=self._run_playback, daemon=True).start()

    def _run_playback(self):
        self.recorder.play_all(self.stop_event)
        self.finish_playback()

    def finish_playback(self):
        if self.interrupt_listener:
            self.interrupt_listener.stop()
            self.interrupt_listener = None
        messagebox.showinfo("Playback", "Macro finished." if not self.stop_event.is_set() else "Macro interrupted.")

    def clear_all(self):
        self.recorder.clear_all()

    def save_macro(self):
        file = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if file:
            self.recorder.save_macro(file)
            messagebox.showinfo("Save", "Macro saved.")

    def load_macro(self):
        file = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if file:
            self.recorder.load_macro(file)


if __name__ == "__main__":
    root = tk.Tk()
    app = MacroEditorApp(root)
    root.mainloop()