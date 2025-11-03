# src/playback_manager.py
import threading
import tkinter.messagebox as messagebox
from pynput import keyboard


class PlaybackManager:
    def __init__(self, app):
        self.app = app
        self.stop_event = None
        self.interrupt_listener = None
        self.pressed = set()

    def start_playback(self):
        self.stop_event = threading.Event()
        self.pressed = set()

        def on_press(k):
            self.pressed.add(k)
            if {keyboard.Key.ctrl, keyboard.Key.alt, keyboard.Key.enter}.issubset(self.pressed):
                self.stop_event.set()

        def on_release(k):
            if k in self.pressed:
                self.pressed.remove(k)

        self.interrupt_listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        self.interrupt_listener.start()
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        self.app.recorder.play_all(self.stop_event)
        self.app.root.after(0, self.finish)

    def finish(self):
        if self.interrupt_listener:
            self.interrupt_listener.stop()
            self.interrupt_listener = None
        msg = "Macro finished." if not self.stop_event.is_set() else "Macro interrupted."
        messagebox.showinfo("Playback", msg)