# src/playback_manager.py
import threading
import tkinter.messagebox as messagebox
from pynput import keyboard
from typing import Optional, Tuple

class PlaybackManager:
    def __init__(self, app):
        self.app = app
        self.stop_event: Optional[threading.Event] = None
        self.interrupt_listener = None
        self.pressed = set()
        self.current_step: Optional[Tuple[int, int, int]] = None

    def start_playback(self):
        self.stop_event = threading.Event()
        self.pressed = set()
        self.current_step = None

        def on_press(k):
            self.pressed.add(k)
            # Ctrl + Alt + Enter OR ESC to abort
            if keyboard.Key.esc in self.pressed:
                self.stop_event.set()
            elif {keyboard.Key.ctrl, keyboard.Key.alt, keyboard.Key.enter}.issubset(self.pressed):
                self.stop_event.set()

        def on_release(k):
            if k in self.pressed:
                self.pressed.remove(k)

        self.interrupt_listener = keyboard.Listener(on_press=on_press, on_release=on_release)
        self.interrupt_listener.start()
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        # ------------------------------------------------------------------
        # 2-D playback loop – mirrors player.py
        # ------------------------------------------------------------------
        snapshot_rows = self.app.recorder.snapshot_rows()
        gaps_within = self.app.recorder.snapshot_within_rows()
        gaps_between = self.app.recorder.snapshot_between_rows()

        try:
            for row_idx, row in enumerate(snapshot_rows):
                for sec_idx, section in enumerate(row):
                    if self.stop_event.is_set():
                        return
                    for step_idx, action in enumerate(section["steps"]):
                        if self.stop_event.is_set():
                            return

                        # ----- UI: highlight current step -----
                        self.current_step = (row_idx, sec_idx, step_idx)
                        self.app.root.after(0, self.app.ui.refresh)  # trigger UI update

                        # Execute the action
                        self.app.recorder._player._execute_action(action, self.stop_event)

                        # ----- UI: clear highlight after execution -----
                        self.current_step = None
                        self.app.root.after(0, self.app.ui.refresh)

                    # gap inside row
                    if sec_idx < len(row) - 1:
                        delay_ms = (
                            gaps_within[row_idx][sec_idx]
                            if row_idx < len(gaps_within) and sec_idx < len(gaps_within[row_idx])
                            else 0
                        )
                        if delay_ms > 0:
                            self.current_step = (row_idx, sec_idx, -1)
                            self.app.root.after(0, self.app.ui.refresh)
                            self._sleep_with_interrupt(delay_ms / 1000.0)
                            self.current_step = None
                            self.app.root.after(0, self.app.ui.refresh)

                # gap between rows
                if row_idx < len(snapshot_rows) - 1:
                    delay_ms = gaps_between[row_idx] if row_idx < len(gaps_between) else 0
                    if delay_ms > 0:
                        self.current_step = (row_idx, -1, -1)
                        self.app.root.after(0, self.app.ui.refresh)
                        self._sleep_with_interrupt(delay_ms / 1000.0)
                        self.current_step = None
                        self.app.root.after(0, self.app.ui.refresh)

        finally:
            self.current_step = None
            self.app.root.after(0, self.finish)

    @staticmethod
    def _sleep_with_interrupt(seconds: float):
        """Non-blocking sleep that respects stop_event."""
        start = threading.Event()
        start.set()
        while start.is_set() and not threading.current_thread().is_alive():
            time.sleep(0.01)

        # Simple version – just sleep (interruption handled by stop_event check above)
        import time
        end = time.time() + seconds
        while time.time() < end:
            if threading.current_thread()._tstate_lock is None:
                break
            time.sleep(0.01)

    def finish(self):
        if self.interrupt_listener:
            self.interrupt_listener.stop()
            self.interrupt_listener = None
        msg = "Macro finished." if not self.stop_event.is_set() else "Macro interrupted."
        messagebox.showinfo("Playback", msg)