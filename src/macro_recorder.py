import time
from pynput import keyboard, mouse
import pyautogui
import json
import threading


class MacroRecorderCore:
    def __init__(self):
        self.sections = []
        self.delays_between = []
        self.recording = False
        self.listener = None
        self.mouse_listener = None
        self.last_time = None
        self.pressed_keys = set()
        self.active_section_index = None
        self.ui_callback = None
        self.playback_ui_callback = None
        self._lock = threading.Lock()
        self._last_ui_update = 0
        self._ui_update_interval = 0.1  # 100ms

    def _notify_ui(self):
        current_time = time.time()
        if current_time - self._last_ui_update >= self._ui_update_interval:
            cb = self.ui_callback
            if cb:
                try:
                    cb()
                    self._last_ui_update = current_time
                except Exception:
                    pass

    def _playback_notify(self, section_idx, step_idx, active):
        cb = self.playback_ui_callback
        if cb:
            try:
                cb(section_idx, step_idx, active)
            except Exception:
                pass

    def _ensure_gap_count(self):
        n = max(0, len(self.sections) - 1)
        if len(self.delays_between) < n:
            self.delays_between.extend([0] * (n - len(self.delays_between)))
        elif len(self.delays_between) > n:
            self.delays_between = self.delays_between[:n]

    def start_recording(self, section_index):
        with self._lock:
            if self.recording or section_index is None or section_index < 0 or section_index >= len(self.sections):
                return
            self.recording = True
            self.active_section_index = section_index
            self.pressed_keys.clear()
            self.last_time = time.time() * 1000

            try:
                self.listener = keyboard.Listener(on_press=self._on_press, on_release=self._on_release)
                self.listener.start()
                self.mouse_listener = mouse.Listener(on_click=self._on_mouse_click)
                self.mouse_listener.start()
            except Exception as e:
                self.recording = False
                self.active_section_index = None
                raise e

        self._notify_ui()

    def stop_recording(self):
        with self._lock:
            if not self.recording:
                return
            self.recording = False
            if self.listener:
                self.listener.stop()
                self.listener = None
            if self.mouse_listener:
                self.mouse_listener.stop()
                self.mouse_listener = None
            if self.active_section_index is not None:
                steps = self.sections[self.active_section_index]["steps"]
                if steps and steps[-1].get("type") in ("mouse_press", "mouse_release"):
                    steps.pop()
            self.pressed_keys.clear()
            self.active_section_index = None
        self._notify_ui()

    def _normalize_key(self, key):
        try:
            return key.char
        except AttributeError:
            return str(key).replace("Key.", "")

    def _on_press(self, key):
        with self._lock:
            if self.active_section_index is None:
                return
            current_time = time.time() * 1000
            k = self._normalize_key(key)

            if k not in self.pressed_keys:
                self.pressed_keys.add(k)
                if self.last_time is not None:
                    delay = int(current_time - self.last_time)
                    self._add_step_no_lock({"type": "delay", "delay": delay, "unit": "ms"})
                self._add_step_no_lock({"type": "press", "key": k})
                self.last_time = current_time
        self._notify_ui()

    def _on_release(self, key):
        with self._lock:
            if self.active_section_index is None:
                return
            current_time = time.time() * 1000
            k = self._normalize_key(key)

            if k in self.pressed_keys:
                self.pressed_keys.remove(k)
                if self.last_time is not None:
                    delay = int(current_time - self.last_time)
                    self._add_step_no_lock({"type": "delay", "delay": delay, "unit": "ms"})
                self._add_step_no_lock({"type": "release", "key": k})
                self.last_time = current_time
        self._notify_ui()

    def _on_mouse_click(self, x, y, button, pressed):
        with self._lock:
            if not self.recording or self.active_section_index is None:
                return
            current_time = time.time() * 1000
            button_map = {
                mouse.Button.left: 'left',
                mouse.Button.right: 'right',
                mouse.Button.middle: 'middle'
            }
            button_str = button_map.get(button)
            if button_str is None:
                return
            action_type = "mouse_press" if pressed else "mouse_release"
            if self.last_time is not None:
                delay = int(current_time - self.last_time)
                if delay > 0:
                    self._add_step_no_lock({"type": "delay", "delay": delay, "unit": "ms"})
            self._add_step_no_lock({"type": action_type, "x": int(x), "y": int(y), "button": button_str})
            self.last_time = current_time
        self._notify_ui()

    def add_section(self, name="New Section"):
        with self._lock:
            self.sections.append({"name": name, "steps": []})
            self._ensure_gap_count()
            idx = len(self.sections) - 1
        self._notify_ui()
        return idx

    def rename_section(self, idx, name):
        with self._lock:
            if 0 <= idx < len(self.sections):
                self.sections[idx]["name"] = name
        self._notify_ui()

    def delete_section(self, idx):
        with self._lock:
            if not (0 <= idx < len(self.sections)):
                return
            n = len(self.sections)
            if n == 0:
                return

            if n == 1:
                self.sections.pop(idx)
                self.delays_between.clear()
            else:
                if idx == 0:
                    self.sections.pop(0)
                    if self.delays_between:
                        self.delays_between.pop(0)
                elif idx == n - 1:
                    self.sections.pop()
                    if self.delays_between:
                        self.delays_between.pop()
                else:
                    left = self.delays_between[idx - 1]
                    right = self.delays_between[idx]
                    merged = int(left) + int(right)
                    self.sections.pop(idx)
                    self.delays_between[idx - 1] = merged
                    self.delays_between.pop(idx)
            if self.active_section_index is not None:
                if self.active_section_index == idx:
                    self.active_section_index = None
                elif self.active_section_index > idx:
                    self.active_section_index -= 1

            self._ensure_gap_count()
        self._notify_ui()

    def _add_step_no_lock(self, step):
        if self.active_section_index is None:
            return
        self.sections[self.active_section_index]["steps"].append(step)

    def add_delay_step(self, section_index, delay_ms):
        with self._lock:
            if 0 <= section_index < len(self.sections):
                self.sections[section_index]["steps"].append({"type": "delay", "delay": int(delay_ms), "unit": "ms"})
        self._notify_ui()

    def delete_step(self, section_index, step_index):
        with self._lock:
            if 0 <= section_index < len(self.sections):
                steps = self.sections[section_index]["steps"]
                if 0 <= step_index < len(steps):
                    del steps[step_index]
        self._notify_ui()

    def move_step_up(self, section_index, step_index):
        with self._lock:
            if 0 <= section_index < len(self.sections):
                steps = self.sections[section_index]["steps"]
                if 1 <= step_index < len(steps):
                    steps[step_index - 1], steps[step_index] = steps[step_index], steps[step_index - 1]
        self._notify_ui()

    def move_step_down(self, section_index, step_index):
        with self._lock:
            if 0 <= section_index < len(self.sections):
                steps = self.sections[section_index]["steps"]
                if 0 <= step_index < len(steps) - 1:
                    steps[step_index + 1], steps[step_index] = steps[step_index], steps[step_index + 1]
        self._notify_ui()

    def block_move_up(self, section_index, start_idx, end_idx):
        with self._lock:
            if 0 <= section_index < len(self.sections):
                steps = self.sections[section_index]["steps"]
                if 0 <= start_idx <= end_idx < len(steps) and start_idx > 0:
                    block = steps[start_idx:end_idx + 1]
                    steps[start_idx:end_idx + 1] = []
                    steps[start_idx - 1:start_idx - 1] = block
        self._notify_ui()

    def block_move_down(self, section_index, start_idx, end_idx):
        with self._lock:
            if 0 <= section_index < len(self.sections):
                steps = self.sections[section_index]["steps"]
                if 0 <= start_idx <= end_idx < len(steps) - 1:
                    block = steps[start_idx:end_idx + 1]
                    steps[start_idx:end_idx + 1] = []
                    steps[end_idx + 1:end_idx + 1] = block
        self._notify_ui()

    def edit_delay(self, section_index, step_index, new_delay_ms):
        with self._lock:
            if 0 <= section_index < len(self.sections):
                steps = self.sections[section_index]["steps"]
                if 0 <= step_index < len(steps):
                    step = steps[step_index]
                    if step.get("type") == "delay":
                        step["delay"] = int(new_delay_ms)
                        step["unit"] = "ms"
        self._notify_ui()

    def set_between_delay(self, gap_index, ms):
        with self._lock:
            if 0 <= gap_index < len(self.delays_between):
                self.delays_between[gap_index] = int(ms)
        self._notify_ui()

    def clear_all(self):
        with self._lock:
            self.sections.clear()
            self.delays_between.clear()
            self.active_section_index = None
        self._notify_ui()

    def move_section_left(self, idx):
        with self._lock:
            if 1 <= idx < len(self.sections):
                self.sections[idx - 1], self.sections[idx] = self.sections[idx], self.sections[idx - 1]
                if self.active_section_index == idx:
                    self.active_section_index = idx - 1
                elif self.active_section_index == idx - 1:
                    self.active_section_index = idx
        self._notify_ui()

    def move_section_right(self, idx):
        with self._lock:
            if 0 <= idx < len(self.sections) - 1:
                self.sections[idx + 1], self.sections[idx] = self.sections[idx], self.sections[idx + 1]
                if self.active_section_index == idx:
                    self.active_section_index = idx + 1
                elif self.active_section_index == idx + 1:
                    self.active_section_index = idx
        self._notify_ui()

    def play_all(self, stop_event=None):
        snapshot = self.snapshot_sections()
        gaps = self.snapshot_between_delays()
        for s_idx, section in enumerate(snapshot):
            for a_idx, action in enumerate(section["steps"]):
                if stop_event and stop_event.is_set():
                    return
                self._playback_notify(s_idx, a_idx, True)
                self._execute_action(action, stop_event)
                self._playback_notify(s_idx, a_idx, False)
            if s_idx < len(snapshot) - 1:
                delay_ms = int(gaps[s_idx]) if s_idx < len(gaps) else 0
                if delay_ms > 0:
                    self._playback_notify(s_idx, -1, True)
                    self._sleep_with_interrupt(delay_ms / 1000.0, stop_event)
                    self._playback_notify(s_idx, -1, False)

    def _sleep_with_interrupt(self, seconds, stop_event=None):
        start = time.time()
        while time.time() - start < seconds:
            if stop_event and stop_event.is_set():
                return
            time.sleep(0.01)

    def _execute_action(self, action, stop_event=None):
        t = action.get("type")
        if t == "delay":
            unit = action.get("unit", "ms")
            if unit == "ms":
                sleep_time = action["delay"] / 1000
            elif unit == "secs":
                sleep_time = action["delay"]
            elif unit == "mins":
                sleep_time = action["delay"] * 60
            elif unit == "hrs":
                sleep_time = action["delay"] * 3600
            else:
                sleep_time = action["delay"] / 1000
            self._sleep_with_interrupt(sleep_time, stop_event)
        elif t == "press":
            key = action.get("key")
            if key in ("cmd", "cmd_r", "win"):
                pyautogui.keyDown("winleft")
            else:
                pyautogui.keyDown(key)
        elif t == "release":
            key = action.get("key")
            if key in ("cmd", "cmd_r", "win"):
                pyautogui.keyUp("winleft")
            else:
                pyautogui.keyUp(key)
        elif t == "mouse_press":
            x, y, btn = action["x"], action["y"], action["button"]
            pyautogui.moveTo(x, y)
            pyautogui.mouseDown(button=btn)
        elif t == "mouse_release":
            x, y, btn = action["x"], action["y"], action["button"]
            pyautogui.moveTo(x, y)
            pyautogui.mouseUp(button=btn)

    def save_macro(self, path):
        with self._lock:
            data = {
                "sections": self.sections,
                "delays_between": self.delays_between
            }
        with open(path, "w") as f:
            json.dump(data, f)

    def load_macro(self, path):
        with open(path, "r") as f:
            data = json.load(f)
        with self._lock:
            if isinstance(data, list):
                self.sections = data
                self.delays_between = [0] * max(0, len(self.sections) - 1)
            else:
                self.sections = data.get("sections", [])
                self.delays_between = data.get("delays_between", [0] * max(0, len(self.sections) - 1))
            self._ensure_gap_count()
        self._notify_ui()

    def snapshot_sections(self):
        with self._lock:
            return [{"name": s["name"], "steps": list(s["steps"])} for s in self.sections]

    def snapshot_between_delays(self):
        with self._lock:
            return list(self.delays_between)