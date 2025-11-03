# top_bar.py
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

class TopBar:
    def __init__(self, parent, app):
        self.app = app
        self.frame = tk.Frame(parent)
        self.frame.pack(side="top", fill="x", pady=6)

        self.quick_delay_var = tk.StringVar(value="250")
        self.auto_minimize_var = tk.BooleanVar(value=False)

        self._create_widgets()

    def _create_widgets(self):
        buttons = [
            ("Add Column", self.app.add_section),
            ("Add String", self.app.add_string),
            ("Start Recording", self.app.toggle_recording),
            ("Play Macro", self.app.play_macro),
            ("Save", self.app.save_macro),
            ("Load", self.app.load_macro),
            ("Clear All", self.app.clear_all),
        ]

        for text, command in buttons:
            btn = tk.Button(self.frame, text=text, command=command)
            btn.pack(side="left", padx=4)
            if text == "Start Recording":
                self.record_button = btn

        # Quick delay input
        tk.Label(self.frame, text="Step Delay ms:").pack(side="left", padx=(16, 4))
        tk.Entry(self.frame, textvariable=self.quick_delay_var, width=6).pack(side="left")
        tk.Button(self.frame, text="Add Step Delay to Selected", command=self.app.add_quick_delay).pack(side="left", padx=4)

        # Auto-minimize checkbox
        tk.Checkbutton(
            self.frame,
            text="Auto-minimize when recording",
            variable=self.auto_minimize_var
        ).pack(side="left", padx=8)

    def update_record_button(self, recording):
        if recording:
            self.record_button.config(text="Stop Recording", bg="red")
        else:
            self.record_button.config(text="Start Recording", bg="SystemButtonFace")

    def get_auto_minimize(self):
        return self.auto_minimize_var.get()