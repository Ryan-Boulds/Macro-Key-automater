# src/ui_components.py
import tkinter as tk
from tkinter import messagebox

STEP_WIDTH = 18
STEP_HEIGHT = 2

def step_label(step):
    t = step.get("type")
    if t == "delay":
        return f"Delay {step['delay']} {step.get('unit','ms')}"
    elif t == "press":
        return f"{step['key']} (press)"
    elif t == "release":
        return f"{step['key']} (release)"
    elif t == "mouse_press":
        return f"Mouse {step['button']} press @ ({step['x']}, {step['y']})"
    elif t == "mouse_release":
        return f"Mouse {step['button']} release @ ({step['x']}, {step['y']})"
    elif t == "key_group":
        keys = set()
        hold_ms = 0
        for sub in step["sub_steps"]:
            if sub["type"] == "press":
                keys.add(sub["key"])
            elif sub["type"] == "delay":
                hold_ms += sub["delay"]
        if len(keys) == 1:
            return f"Hold {list(keys)[0]} for {hold_ms}ms" if hold_ms > 0 else f"Tap {list(keys)[0]}"
        else:
            return f"Chord: {' + '.join(sorted(keys))}"
    elif t == "mouse_click":
        btn = step["button"].capitalize()
        pos = f"({step['x']}, {step['y']})"
        release_x = step.get("release_x", step["x"])
        release_y = step.get("release_y", step["y"])
        if release_x != step["x"] or release_y != step["y"]:
            pos += f" to ({release_x}, {release_y})"
        hold = step["hold_ms"]
        if hold > 0:
            return f"{btn} hold {hold}ms at {pos}"
        else:
            return f"{btn} click at {pos}"
    elif t == "typed":
        return f"typed: \"{step['chars']}\""
    elif t == "mouse_drag":
        return f"Drag {step['button']} from ({step['start_x']}, {step['start_y']}) to ({step['end_x']}, {step['end_y']}) in {step['duration_ms']}ms"
    return "Unknown"


def render_gap_chip(app, gap_index, value_ms, parent=None):
    if parent is None:
        parent = app.main_frame

    frame = tk.Frame(parent, width=60)
    frame.pack_propagate(False)

    chip = tk.Frame(frame, bd=1, relief="ridge", bg="white")
    chip.pack(fill="both", expand=True, padx=2, pady=2)

    row_idx = gap_index[0]
    while len(app.gap_chips) <= row_idx:
        app.gap_chips.append([])
    app.gap_chips[row_idx].append(chip)

    tk.Label(chip, text="Between", font=("TkDefaultFont", 8)).pack(pady=(6, 0))
    var = tk.StringVar(value=str(value_ms))
    entry = tk.Entry(chip, textvariable=var, width=6, justify="center")
    entry.pack(pady=2)

    def apply():
        try:
            ms = int(float(var.get()))
            if ms < 0:
                raise ValueError
            app.recorder.set_gap_delay(gap_index, ms)
        except ValueError:
            messagebox.showerror("Error", "Enter a valid delay (ms).")

    tk.Button(chip, text="Set ms", command=apply).pack(pady=(0, 6))
    return frame


def render_section(app, idx, section, parent=None):
    if parent is None:
        parent = app.main_frame

    is_active = (idx == app.active_section_index)
    border_color = "#0078D7" if is_active else "#cccccc"

    frame = tk.Frame(parent, bd=2, relief="groove", highlightthickness=2)
    frame.configure(highlightbackground=border_color, highlightcolor=border_color)

    # Header
    header = tk.Frame(frame)
    header.pack(fill="x", padx=6, pady=6)

    name_var = tk.StringVar(value=section["name"])
    name_entry = tk.Entry(header, textvariable=name_var, justify="center")
    name_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
    name_entry.bind("<Return>", lambda _e, i=idx, v=name_var: app.recorder.rename_section(i, v.get()))
    name_entry.bind("<FocusOut>", lambda _e, i=idx, v=name_var: app.recorder.rename_section(i, v.get()))

    tk.Button(header, text="Left", width=3, command=lambda i=idx: app.move_section_left(i)).pack(side="left", padx=2)
    tk.Button(header, text="Right", width=3, command=lambda i=idx: app.move_section_right(i)).pack(side="left", padx=2)
    record_btn = tk.Button(header, text="Record Here", command=lambda i=idx: app.select_section(i))
    if is_active:
        record_btn.config(bg="red")
    record_btn.pack(side="left", padx=6)
    tk.Button(header, text="Delete", command=lambda i=idx: app.delete_section(i)).pack(side="left", padx=6)

    # Steps
    steps_wrap = tk.Frame(frame)
    steps_wrap.pack(fill="both", expand=True, padx=6, pady=(0, 6))

    row_idx, sec_idx = idx

    # Ensure structures
    while len(app.step_labels) <= row_idx:
        app.step_labels.append([])
    while len(app.step_labels[row_idx]) <= sec_idx:
        app.step_labels[row_idx].append([])
    app.step_labels[row_idx][sec_idx] = []

    # Ensure step_menus exists
    if not hasattr(app, 'step_menus'):
        app.step_menus = []

    for s_idx, step in enumerate(section["steps"]):
        step_frame = tk.Frame(steps_wrap)
        step_frame.pack(fill="x", pady=1)

        text = step_label(step)
        bg_color = "white"
        if app.last_recorded_step == (row_idx, sec_idx, s_idx):
            bg_color = "#FFFF99"
        elif ((row_idx, sec_idx), s_idx) in app.selection.selected_indices:
            bg_color = "#D3D3D3"

        lbl = tk.Label(
            step_frame, text=text, bd=1, relief="solid",
            width=STEP_WIDTH, height=STEP_HEIGHT,
            anchor="w", justify="left", bg=bg_color, wraplength=300
        )
        lbl.pack(side="left", fill="x", expand=True)
        app.step_labels[row_idx][sec_idx].append(lbl)

        # Selection
        def make_toggle(si, sti):
            def toggle(event):
                app.selection.toggle(si, sti, lbl, event.state & 0x4)
            return toggle
        lbl.bind("<Button-1>", make_toggle((row_idx, sec_idx), s_idx))

        # Context menu
        menu = tk.Menu(app.root, tearoff=0)
        app.step_menus.append(menu)
        menu.add_command(label="Delete", command=lambda si=(row_idx, sec_idx), sti=s_idx: app.delete_step(si, sti))
        if step.get("type") == "delay":
            menu.add_command(label="Edit Delay...", command=lambda: messagebox.showinfo("Edit", "Not implemented yet"))

        def show_menu(e, m=menu):
            try:
                m.tk_popup(e.x_root, e.y_root)
            finally:
                m.grab_release()
        lbl.bind("<Button-3>", show_menu)

    return frame


# === FULL DIALOGS (exactly as you wrote them) ===
def add_typed_dialog(app, section_idx):
    dialog = tk.Toplevel(app.root)
    dialog.title("Add Typed String")
    frame = tk.Frame(dialog)
    frame.pack(pady=10, padx=10)
    tk.Label(frame, text="String:").grid(row=0, column=0)
    chars_var = tk.StringVar(value="")
    tk.Entry(frame, textvariable=chars_var, width=30).grid(row=0, column=1)
    tk.Label(frame, text="Delay between Keypresses (ms):").grid(row=1, column=0)
    delay_var = tk.StringVar(value="15")
    tk.Entry(frame, textvariable=delay_var, width=10).grid(row=1, column=1)
    def save():
        try:
            chars = chars_var.get()
            delay = int(delay_var.get())
            if not chars:
                messagebox.showerror("Error", "String cannot be empty.")
                return
            if delay < 0:
                messagebox.showerror("Error", "Delay must be non-negative.")
                return
            app.recorder.add_typed_step(section_idx, chars, delay)
        except ValueError:
            messagebox.showerror("Error", "Invalid delay.")
            return
        app.render_sections()
        dialog.destroy()
    tk.Button(dialog, text="Save", command=save).pack(pady=10)


def edit_key_group(app, section_idx, step_idx):
    sections = app.recorder.snapshot_sections()
    step = sections[section_idx]["steps"][step_idx]
    if step.get("type") != "key_group":
        return
    sub_steps = step["sub_steps"]
    dialog = tk.Toplevel(app.root)
    dialog.title("Edit Key Group")
    frame = tk.Frame(dialog)
    frame.pack(pady=10, padx=10)

    chars = []
    modifier_keys = {"cmd", "cmd_r", "win", "ctrl", "alt", "shift", "enter", "tab", "esc", "backspace", "delete"}
    shift_pressed = False
    current_char = None
    shift_map = {
        "1": "!", "2": "@", "3": "#", "4": "$", "5": "%", "6": "^", "7": "&", "8": "*", "9": "(", "0": ")",
        "-": "_", "=": "+", ".": "."
    }
    for sub in sub_steps:
        if sub["type"] == "press":
            if sub["key"] == "shift":
                shift_pressed = True
            elif shift_pressed and sub["key"] in shift_map:
                chars.append(shift_map[sub["key"]])
                current_char = shift_map[sub["key"]]
            elif sub["key"] == "space":
                chars.append(" ")
                current_char = " "
            elif sub["key"] not in modifier_keys:
                chars.append(sub["key"])
                current_char = sub["key"]
        elif sub["type"] == "release" and (sub["key"] == current_char or sub["key"] == "shift" or sub["key"] == "space" or (shift_pressed and sub["key"] in shift_map)):
            if sub["key"] == "shift":
                shift_pressed = False
            current_char = None

    entries = []
    for idx, sub in enumerate(sub_steps):
        if sub["type"] == "delay":
            tk.Label(frame, text="Delay:").grid(row=idx, column=0)
            var = tk.StringVar(value=str(sub["delay"]))
            entry = tk.Entry(frame, textvariable=var)
            entry.grid(row=idx, column=1)
            entries.append((idx, var))
        else:
            action = "Press" if sub["type"] == "press" else "Release"
            tk.Label(frame, text=f"{action} {sub['key']}").grid(row=idx, column=0, columnspan=2)

    tk.Label(frame, text="Convert to String (optional):").grid(row=len(sub_steps), column=0)
    string_var = tk.StringVar(value="".join(chars) if chars else "")
    tk.Entry(frame, textvariable=string_var, width=30).grid(row=len(sub_steps), column=1)

    tk.Label(frame, text="Delay between Keypresses (ms):").grid(row=len(sub_steps)+1, column=0)
    delay_var = tk.StringVar(value="15")
    tk.Entry(frame, textvariable=delay_var, width=10).grid(row=len(sub_steps)+1, column=1)

    def save():
        try:
            new_string = string_var.get()
            delay = int(delay_var.get())
            if delay < 0:
                messagebox.showerror("Error", "Delay must be non-negative.")
                return
            for i, var in entries:
                try:
                    sub_steps[i]["delay"] = int(var.get())
                except ValueError:
                    messagebox.showerror("Error", "Invalid delay.")
                    return
            with app.recorder._lock:
                if new_string:
                    app.recorder.sections[section_idx]["steps"][step_idx] = {
                        "type": "typed",
                        "chars": new_string,
                        "delays": [delay] * (len(new_string) - 1)
                    }
                else:
                    app.recorder.sections[section_idx]["steps"][step_idx]["sub_steps"] = sub_steps
            app.render_sections()
            dialog.destroy()
        except ValueError:
            messagebox.showerror("Error", "Invalid delay.")
            return
    tk.Button(dialog, text="Save", command=save).pack(pady=10)


def edit_mouse_click(app, section_idx, step_idx):
    sections = app.recorder.snapshot_sections()
    step = sections[section_idx]["steps"][step_idx]
    if step.get("type") != "mouse_click":
        return
    dialog = tk.Toplevel(app.root)
    dialog.title("Edit Mouse Click")
    frame = tk.Frame(dialog)
    frame.pack(pady=10, padx=10)
    tk.Label(frame, text="Button:").grid(row=0, column=0)
    tk.Label(frame, text=step["button"]).grid(row=0, column=1)
    tk.Label(frame, text="Press X:").grid(row=1, column=0)
    x_var = tk.StringVar(value=str(step["x"]))
    tk.Entry(frame, textvariable=x_var).grid(row=1, column=1)
    tk.Label(frame, text="Press Y:").grid(row=2, column=0)
    y_var = tk.StringVar(value=str(step["y"]))
    tk.Entry(frame, textvariable=y_var).grid(row=2, column=1)
    tk.Label(frame, text="Hold ms:").grid(row=3, column=0)
    hold_var = tk.StringVar(value=str(step["hold_ms"]))
    tk.Entry(frame, textvariable=hold_var).grid(row=3, column=1)
    tk.Label(frame, text="Release X:").grid(row=4, column=0)
    rel_x_var = tk.StringVar(value=str(step.get("release_x", step["x"])))
    tk.Entry(frame, textvariable=rel_x_var).grid(row=4, column=1)
    tk.Label(frame, text="Release Y:").grid(row=5, column=0)
    rel_y_var = tk.StringVar(value=str(step.get("release_y", step["y"])))
    tk.Entry(frame, textvariable=rel_y_var).grid(row=5, column=1)
    def save():
        try:
            step["x"] = int(x_var.get())
            step["y"] = int(y_var.get())
            step["hold_ms"] = int(hold_var.get())
            step["release_x"] = int(rel_x_var.get())
            step["release_y"] = int(rel_y_var.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid input.")
            return
        with app.recorder._lock:
            app.recorder.sections[section_idx]["steps"][step_idx] = step
        app.render_sections()
        dialog.destroy()
    tk.Button(dialog, text="Save", command=save).pack(pady=10)


def edit_typed(app, section_idx, step_idx):
    sections = app.recorder.snapshot_sections()
    step = sections[section_idx]["steps"][step_idx]
    if step.get("type") != "typed":
        return
    dialog = tk.Toplevel(app.root)
    dialog.title("Edit Typed String")
    frame = tk.Frame(dialog)
    frame.pack(pady=10, padx=10)
    tk.Label(frame, text=f"Entered String: {step['chars']}").grid(row=0, column=0, columnspan=2)
    tk.Label(frame, text="New String:").grid(row=1, column=0)
    chars_var = tk.StringVar(value=step["chars"])
    tk.Entry(frame, textvariable=chars_var, width=30).grid(row=1, column=1)
    tk.Label(frame, text="Delay between Keypresses (ms):").grid(row=2, column=0)
    delay_var = tk.StringVar(value=str(step["delays"][0] if step["delays"] else 15))
    tk.Entry(frame, textvariable=delay_var, width=10).grid(row=2, column=1)
    def save():
        try:
            chars = chars_var.get()
            delay = int(delay_var.get())
            if not chars:
                messagebox.showerror("Error", "String cannot be empty.")
                return
            if delay < 0:
                messagebox.showerror("Error", "Delay must be non-negative.")
                return
            step["chars"] = chars
            step["delays"] = [delay] * (len(chars) - 1)
        except ValueError:
            messagebox.showerror("Error", "Invalid delay.")
            return
        with app.recorder._lock:
            app.recorder.sections[section_idx]["steps"][step_idx] = step
        app.render_sections()
        dialog.destroy()
    tk.Button(dialog, text="Save", command=save).pack(pady=10)