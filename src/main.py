import tkinter as tk
from tkinter import filedialog, messagebox
import json
import time
from pynput import keyboard
import pyautogui
import threading

class MacroRecorder:
    def __init__(self, root):
        self.root = root
        self.root.title("Macro Recorder")
        self.root.geometry("600x400")
        self.recording = False
        self.macro = []
        self.listener = None
        self.last_time = None
        self.pressed_keys = set()
        self.action_windows = []
        self.max_columns = 1
        self.action_width = 500
        self.action_height = 50
        self.pad = 5

        # Main frame
        self.main_frame = tk.Frame(root)
        self.main_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        # Canvas + Scrollbar
        self.canvas = tk.Canvas(self.main_frame)
        self.scrollbar = tk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        def on_mouse_wheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self.canvas.bind_all("<MouseWheel>", on_mouse_wheel)

        # Buttons
        btn_frame = tk.Frame(self.main_frame)
        btn_frame.pack(pady=5)

        self.start_btn = tk.Button(btn_frame, text="Start Recording", command=self.start_recording)
        self.start_btn.grid(row=0, column=0, padx=5)

        self.stop_btn = tk.Button(btn_frame, text="Stop Recording", command=self.stop_recording, state=tk.DISABLED)
        self.stop_btn.grid(row=0, column=1, padx=5)

        self.play_btn = tk.Button(btn_frame, text="Play Macro", command=self.play_macro)
        self.play_btn.grid(row=0, column=2, padx=5)

        self.save_btn = tk.Button(btn_frame, text="Save Macro", command=self.save_macro)
        self.save_btn.grid(row=0, column=3, padx=5)

        self.load_btn = tk.Button(btn_frame, text="Load Macro", command=self.load_macro)
        self.load_btn.grid(row=0, column=4, padx=5)

        self.clear_btn = tk.Button(btn_frame, text="Clear All", command=self.clear_all)
        self.clear_btn.grid(row=0, column=5, padx=5)

        # Delay entry
        delay_frame = tk.Frame(self.main_frame)
        delay_frame.pack(pady=5)
        tk.Label(delay_frame, text="Insert Delay (ms):").grid(row=0, column=0)
        self.delay_entry = tk.Entry(delay_frame, width=10)
        self.delay_entry.grid(row=0, column=1)
        tk.Button(delay_frame, text="Add Delay", command=self.add_delay).grid(row=0, column=2, padx=5)
        tk.Label(delay_frame, text="Edit delays in action frame.").grid(row=1, column=0, columnspan=3)

    def on_press(self, key):
        current_time = time.time() * 1000
        try:
            k = key.char
        except AttributeError:
            k = str(key).replace("Key.", "")

        if k not in self.pressed_keys:
            self.pressed_keys.add(k)
            if self.last_time is not None:
                delay = int(current_time - self.last_time)
                self.macro.append({'type': 'delay', 'delay': delay, 'unit': 'ms'})
                self.add_action_to_frame("delay", delay)
            self.macro.append({'type': 'press', 'key': k})
            self.add_action_to_frame("press", k)
            self.last_time = current_time

    def on_release(self, key):
        current_time = time.time() * 1000
        try:
            k = key.char
        except AttributeError:
            k = str(key).replace("Key.", "")

        if k in self.pressed_keys:
            self.pressed_keys.remove(k)
            if self.last_time is not None:
                delay = int(current_time - self.last_time)
                self.macro.append({'type': 'delay', 'delay': delay, 'unit': 'ms'})
                self.add_action_to_frame("delay", delay)
            self.macro.append({'type': 'release', 'key': k})
            self.add_action_to_frame("release", k)
            self.last_time = current_time

    def get_index(self, frame):
        for i, w in enumerate(self.action_windows):
            if w['frame'] == frame:
                return i
        return -1

    def add_action_to_frame(self, action_type, value):
        frame = tk.Frame(self.canvas, bd=1, relief=tk.SOLID, width=self.action_width, height=self.action_height)
        frame.pack_propagate(False)

        content_frame = tk.Frame(frame)
        content_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        idx = len(self.macro) - 1
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Delete", command=lambda f=frame: self.delete_action(self.get_index(f)))

        if action_type == "delay":
            unit = self.macro[-1].get('unit', 'ms') if self.macro else 'ms'
            label = tk.Label(content_frame, text=f"Delay ({unit})", font=("Arial", 10, "bold"))
            label.pack(pady=2)
            delay_entry = tk.Entry(content_frame, width=10)
            delay_entry.insert(0, str(value))
            delay_entry.pack(pady=2)
            delay_entry.bind("<KeyRelease>", lambda e, f=frame: self.update_delay(e, self.get_index(f)))

            menu.add_separator()
            menu.add_command(label="ms", command=lambda f=frame: self.change_unit('ms', self.get_index(f)))
            menu.add_command(label="secs", command=lambda f=frame: self.change_unit('secs', self.get_index(f)))
            menu.add_command(label="mins", command=lambda f=frame: self.change_unit('mins', self.get_index(f)))
            menu.add_command(label="hrs", command=lambda f=frame: self.change_unit('hrs', self.get_index(f)))
        else:
            if action_type == "press":
                label_text = f"{value} (pressed)"
            else:
                label_text = f"{value} (released)"
            label = tk.Label(content_frame, text=label_text, font=("Arial", 10, "bold"))
            label.pack(expand=True, fill=tk.BOTH)
            label.bind("<Button-3>", lambda e: menu.post(e.x_root, e.y_root))

        # Reorder buttons
        arrow_frame = tk.Frame(frame)
        arrow_frame.pack(side=tk.RIGHT, padx=5)
        tk.Button(arrow_frame, text="↑", command=lambda f=frame: self.move_up(self.get_index(f))).pack()
        tk.Button(arrow_frame, text="↓", command=lambda f=frame: self.move_down(self.get_index(f))).pack()

        frame.bind("<Button-3>", lambda e: menu.post(e.x_root, e.y_root))
        content_frame.bind("<Button-3>", lambda e: menu.post(e.x_root, e.y_root))

        row = len(self.action_windows) // self.max_columns
        col = len(self.action_windows) % self.max_columns
        x = col * (self.action_width + self.pad) + self.pad
        y = row * (self.action_height + self.pad) + self.pad
        win_id = self.canvas.create_window(x, y, window=frame, anchor="nw")
        self.action_windows.append({'frame': frame, 'id': win_id})
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def move_up(self, index):
        if index > 0:
            self.macro[index], self.macro[index - 1] = self.macro[index - 1], self.macro[index]
            self.action_windows[index], self.action_windows[index - 1] = self.action_windows[index - 1], self.action_windows[index]
            self.reposition_actions()

    def move_down(self, index):
        if index < len(self.action_windows) - 1:
            self.macro[index], self.macro[index + 1] = self.macro[index + 1], self.macro[index]
            self.action_windows[index], self.action_windows[index + 1] = self.action_windows[index + 1], self.action_windows[index]
            self.reposition_actions()

    def reposition_actions(self):
        for i, w in enumerate(self.action_windows):
            row = i // self.max_columns
            col = i % self.max_columns
            x = col * (self.action_width + self.pad) + self.pad
            y = row * (self.action_height + self.pad) + self.pad
            self.canvas.coords(w['id'], x, y)
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def clear_all(self):
        self.macro.clear()
        for w in self.action_windows:
            self.canvas.delete(w['id'])
            w['frame'].destroy()
        self.action_windows.clear()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def start_recording(self):
        if self.recording:
            return
        self.recording = True
        self.pressed_keys.clear()
        self.last_time = time.time() * 1000

        self.listener = keyboard.Listener(on_press=self.on_press, on_release=self.on_release)
        self.listener.start()

        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        messagebox.showinfo("Recording", "Recording started.\nYou can now switch to other apps.")

    def stop_recording(self):
        if not self.recording:
            return
        self.recording = False
        if self.listener:
            self.listener.stop()
        self.pressed_keys.clear()
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        messagebox.showinfo("Recording", "Recording stopped.")

    def add_delay(self):
        try:
            delay = int(float(self.delay_entry.get()))
            self.macro.append({'type': 'delay', 'delay': delay, 'unit': 'ms'})
            self.add_action_to_frame("delay", delay)
            self.delay_entry.delete(0, tk.END)
        except ValueError:
            messagebox.showerror("Error", "Enter a valid number for delay.")

    def play_macro(self):
        if not self.macro:
            messagebox.showerror("Error", "No macro to play.")
            return
        self.original_bg = self.root.cget("bg")
        self.root.config(bg="white")
        self.main_frame.pack_forget()
        self.play_label = tk.Label(self.root, text="Press Enter to begin macro.\nPress ctrl+alt+enter to stop.", bg="white", fg="black", font=("Arial", 16))
        self.play_label.pack(expand=True, fill=tk.BOTH)
        self.root.bind("<Return>", self.start_playback)
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

    def start_playback(self, event):
        self.play_label.config(text="Running macro...\nPress ctrl+alt+enter to stop.")
        self.root.unbind("<Return>")
        threading.Thread(target=self.playback).start()

    def playback(self):
        for action in self.macro:
            if self.stop_event.is_set():
                break
            if action['type'] == 'delay':
                unit = action.get('unit', 'ms')
                if unit == 'ms':
                    sleep_time = action['delay'] / 1000
                elif unit == 'secs':
                    sleep_time = action['delay']
                elif unit == 'mins':
                    sleep_time = action['delay'] * 60
                elif unit == 'hrs':
                    sleep_time = action['delay'] * 3600
                else:
                    sleep_time = action['delay'] / 1000
                start = time.time()
                while time.time() - start < sleep_time:
                    if self.stop_event.is_set():
                        self.finish_playback()
                        return
                    time.sleep(0.01)
            elif action['type'] == 'press':
                if action['key'] in ['cmd', 'cmd_r', 'win']:
                    pyautogui.keyDown('winleft')
                else:
                    pyautogui.keyDown(action['key'])
            elif action['type'] == 'release':
                if action['key'] in ['cmd', 'cmd_r', 'win']:
                    pyautogui.keyUp('winleft')
                else:
                    pyautogui.keyUp(action['key'])
        self.finish_playback()

    def finish_playback(self):
        self.interrupt_listener.stop()
        self.play_label.pack_forget()
        self.main_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        self.root.config(bg=self.original_bg)
        messagebox.showinfo("Playback", "Macro finished." if not self.stop_event.is_set() else "Macro interrupted.")

    def save_macro(self):
        if not self.macro:
            messagebox.showerror("Error", "No macro to save.")
            return
        file = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if file:
            with open(file, 'w') as f:
                json.dump(self.macro, f)
            messagebox.showinfo("Save", "Macro saved.")

    def load_macro(self):
        file = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if file:
            with open(file, 'r') as f:
                self.macro = json.load(f)
            for action in self.macro:
                if action['type'] == 'delay':
                    action['delay'] = int(action['delay'])
            for w in self.action_windows:
                self.canvas.delete(w['id'])
                w['frame'].destroy()
            self.action_windows.clear()
            for action in self.macro:
                if action['type'] == 'delay':
                    self.add_action_to_frame('delay', action['delay'])
                else:
                    self.add_action_to_frame(action['type'], action['key'])
            messagebox.showinfo("Load", "Macro loaded.")

if __name__ == "__main__":
    root = tk.Tk()
    app = MacroRecorder(root)
    root.mainloop()
