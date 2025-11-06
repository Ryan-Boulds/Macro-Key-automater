# macro_recorder/recorder.py
import time
from pynput import keyboard, mouse
from typing import Set, Optional, Tuple

from .utils import normalize_key
from .merger import merge_steps


class Recorder:
    def __init__(self, core):
        self.core = core
        self.listener: Optional[keyboard.Listener] = None
        self.mouse_listener: Optional[mouse.Listener] = None
        self.last_time: Optional[float] = None
        self.pressed_keys: Set[str] = set()
        self.pressed_mouse_buttons: Set[str] = set()
        self._pre_record_len: int = 0
        self._ui_update_scheduled = False

    def start(self, section_index: Tuple[int, int]):
        row, col = section_index
        with self.core._lock:
            if self.core.recording:
                return
            if not (0 <= row < len(self.core.rows) and 0 <= col < len(self.core.rows[row])):
                return

            self.core.recording = True
            self.core.active_section_index = section_index
            steps = self.core.rows[row][col]["steps"]
            self._pre_record_len = len(steps)
            self.pressed_keys.clear()
            self.pressed_mouse_buttons.clear()
            self.last_time = time.time() * 1000
            self._ui_update_scheduled = False

            try:
                self.listener = keyboard.Listener(
                    on_press=self._on_press,
                    on_release=self._on_release,
                )
                self.listener.start()

                self.mouse_listener = mouse.Listener(
                    on_click=self._on_mouse_click,
                    on_move=self._on_mouse_move,
                )
                self.mouse_listener.start()
            except Exception as e:
                self.core.recording = False
                self.core.active_section_index = None
                raise e

    def stop(self):
        with self.core._lock:
            if not self.core.recording:
                return

            self.core.recording = False
            if self.listener:
                self.listener.stop()
                self.listener = None
            if self.mouse_listener:
                self.mouse_listener.stop()
                self.mouse_listener = None

            if self.core.active_section_index is not None:
                row, col = self.core.active_section_index
                steps = self.core.rows[row][col]["steps"]

                while steps and steps[-1].get("type") in ("mouse_press", "mouse_release", "press", "release"):
                    steps.pop()
                new_part = merge_steps(steps[self._pre_record_len:])
                steps[self._pre_record_len:] = new_part

            self.pressed_keys.clear()
            self.pressed_mouse_buttons.clear()
            self.core.active_section_index = None
            self._ui_update_scheduled = False
        # end lock
        self.core._notify_ui()

    def _schedule_ui_update(self):
        """Throttle UI updates to at most ~30 Hz while recording."""
        if not self._ui_update_scheduled and self.core.ui_callback:
            self._ui_update_scheduled = True
            self.core.root.after(30, self._do_ui_update)

    def _do_ui_update(self):
        self._ui_update_scheduled = False
        self.core._notify_ui()

    def _add_step(self, step: dict):
        if self.core.active_section_index is None:
            return
        row, col = self.core.active_section_index
        self.core.rows[row][col]["steps"].append(step)

    def _on_press(self, key):
        with self.core._lock:
            now = time.time() * 1000
            k = normalize_key(key)
            if k not in self.pressed_keys:
                self.pressed_keys.add(k)
                if self.last_time is not None:
                    delay = int(now - self.last_time)
                    if delay > 0:
                        self._add_step({"type": "delay", "delay": delay, "unit": "ms"})
                self._add_step({"type": "press", "key": k})
                self.last_time = now
        self._schedule_ui_update()

    def _on_release(self, key):
        with self.core._lock:
            now = time.time() * 1000
            k = normalize_key(key)
            if k in self.pressed_keys:
                self.pressed_keys.remove(k)
                if self.last_time is not None:
                    delay = int(now - self.last_time)
                    if delay > 0:
                        self._add_step({"type": "delay", "delay": delay, "unit": "ms"})
                self._add_step({"type": "release", "key": k})
                self.last_time = now
        self._schedule_ui_update()

    def _on_mouse_click(self, x, y, button, pressed):
        with self.core._lock:
            if not self.core.recording or self.core.active_section_index is None:
                return
            now = time.time() * 1000
            button_map = {
                mouse.Button.left: "left",
                mouse.Button.right: "right",
                mouse.Button.middle: "middle",
            }
            button_str = button_map.get(button)
            if button_str is None:
                return

            if pressed:
                self.pressed_mouse_buttons.add(button_str)
            else:
                self.pressed_mouse_buttons.discard(button_str)

            action_type = "mouse_press" if pressed else "mouse_release"
            if self.last_time is not None:
                delay = int(now - self.last_time)
                if delay > 0:
                    self._add_step({"type": "delay", "delay": delay, "unit": "ms"})
            self._add_step(
                {"type": action_type, "x": int(x), "y": int(y), "button": button_str}
            )
            self.last_time = now
        self._schedule_ui_update()

    def _on_mouse_move(self, x, y):
        with self.core._lock:
            if not self.core.recording or self.core.active_section_index is None:
                return
            if not self.pressed_mouse_buttons:  # Only record if mouse button is held
                return

            now = time.time() * 1000
            if self.last_time is not None:
                delay = int(now - self.last_time)
                if delay > 0:
                    self._add_step({"type": "delay", "delay": delay, "unit": "ms"})
            self._add_step({
                "type": "mouse_move",
                "x": int(x),
                "y": int(y),
            })
            self.last_time = now
        self._schedule_ui_update()