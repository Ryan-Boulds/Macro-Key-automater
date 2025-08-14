import time
from pynput import keyboard
import pyautogui
import json
import threading


class MacroRecorderCore:
    """
    Headless core: manages sections, inter-column delays, recording, and playback.
    UI registers `ui_callback` which is invoked after any change.
    """
    def __init__(self):
        # Sections: [{ "name": str, "steps": [ {"type": "press"/"release"/"delay", ...}, ... ] }]
        self.sections = []
        # Delays BETWEEN columns (gaps). If there are N sections, there are N-1 gaps.
        # delays_between[i] is the delay between sections[i] and sections[i+1], in ms.
        self.delays_between = []

        self.recording = False
        self.listener = None
        self.last_time = None
        self.pressed_keys = set()
        self.active_section_index = None

        self.ui_callback = None   # set by UI: a zero-arg callable
        self._lock = threading.Lock()

    # ---------- Internal helpers ----------
    def _notify_ui(self):
        cb = self.ui_callback
        if cb:
            try:
                cb()
            except Exception:
                pass

    def _ensure_gap_count(self):
        """Keep len(delays_between) == max(0, len(sections)-1)."""
        n = max(0, len(self.sections) - 1)
        if len(self.delays_between) < n:
            self.delays_between.extend([0] * (n - len(self.delays_between)))
        elif len(self.delays_between) > n:
            self.delays_between = self.delays_between[:n]

    # ---------- Recording ----------
    def start_recording(self, section_index):
        """Start recording into a given section index (append-only)."""
        with self._lock:
            if self.recording or section_index is None or section_index < 0 or section_index >= len(self.sections):
                return
            self.recording = True
            self.active_section_index = section_index
            self.pressed_keys.clear()
            self.last_time = time.time() * 1000

            self.listener = keyboard.Listener(on_press=self._on_press, on_release=self._on_release)
            self.listener.start()

        self._notify_ui()

    def stop_recording(self):
        with self._lock:
            if not self.recording:
                return
            self.recording = False
            if self.listener:
                self.listener.stop()
                self.listener = None
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

    # ---------- Sections & Steps ----------
    def add_section(self, name="New Section"):
        with self._lock:
            self.sections.append({"name": name, "steps": []})
            # Adding a section increases gaps by 1 if there is at least one prior section.
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
        """Delete a whole section/column and reconcile the BETWEEN gaps.

        If deleting at the ends, remove the adjacent gap.
        If deleting a middle section, merge the two adjacent gaps by SUM
        so the overall time between the neighbors is preserved.
        """
        with self._lock:
            if not (0 <= idx < len(self.sections)):
                return
            n = len(self.sections)
            if n == 0:
                return

            # Reconcile gaps
            if n == 1:
                self.sections.pop(idx)
                self.delays_between.clear()
            else:
                # There are n-1 gaps currently
                if idx == 0:
                    # Remove first section: remove gap[0]
                    self.sections.pop(0)
                    if self.delays_between:
                        self.delays_between.pop(0)
                elif idx == n - 1:
                    # Remove last section: remove last gap
                    self.sections.pop()
                    if self.delays_between:
                        self.delays_between.pop()
                else:
                    # Middle: merge gaps idx-1 and idx into one (sum)
                    left = self.delays_between[idx - 1]
                    right = self.delays_between[idx]
                    merged = int(left) + int(right)
                    self.sections.pop(idx)
                    # Replace left gap with merged, then remove right gap
                    self.delays_between[idx - 1] = merged
                    self.delays_between.pop(idx)
            # Active recording index may shift
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
        """Set delay between sections[gap_index] and sections[gap_index+1]."""
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
        """Swap section idx with idx-1. Gaps remain positioned BETWEEN columns."""
        with self._lock:
            if 1 <= idx < len(self.sections):
                self.sections[idx - 1], self.sections[idx] = self.sections[idx], self.sections[idx - 1]
                # Gaps tied to positions: leave self.delays_between unchanged.
                if self.active_section_index == idx:
                    self.active_section_index = idx - 1
                elif self.active_section_index == idx - 1:
                    self.active_section_index = idx
        self._notify_ui()

    def move_section_right(self, idx):
        with self._lock:
            if 0 <= idx < len(self.sections) - 1:
                self.sections[idx + 1], self.sections[idx] = self.sections[idx], self.sections[idx + 1]
                # Gaps tied to positions: leave self.delays_between unchanged.
                if self.active_section_index == idx:
                    self.active_section_index = idx + 1
                elif self.active_section_index == idx + 1:
                    self.active_section_index = idx
        self._notify_ui()

    # ---------- Playback ----------
    def play_all(self, stop_event=None):
        # iterate left-to-right sections, top-to-bottom steps
        snapshot = self.snapshot_sections()
        gaps = self.snapshot_between_delays()
        for s_idx, section in enumerate(snapshot):
            for action in section["steps"]:
                if stop_event and stop_event.is_set():
                    return
                self._execute_action(action, stop_event)
            # Inter-column delay (between sections)
            if s_idx < len(snapshot) - 1:
                delay_ms = int(gaps[s_idx]) if s_idx < len(gaps) else 0
                self._sleep_with_interrupt(delay_ms / 1000.0, stop_event)

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
                pyautogui.keyDown("winleft")  # Windows key
            else:
                pyautogui.keyDown(key)
        elif t == "release":
            key = action.get("key")
            if key in ("cmd", "cmd_r", "win"):
                pyautogui.keyUp("winleft")
            else:
                pyautogui.keyUp(key)

    # ---------- Persistence ----------
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
            # Backward compatibility with old flat list files
            if isinstance(data, list):
                self.sections = data
                # infer gaps
                self.delays_between = [0] * max(0, len(self.sections) - 1)
            else:
                self.sections = data.get("sections", [])
                self.delays_between = data.get("delays_between", [0] * max(0, len(self.sections) - 1))
            self._ensure_gap_count()
        self._notify_ui()

    # ---------- Snapshots ----------
    def snapshot_sections(self):
        with self._lock:
            return [{"name": s["name"], "steps": list(s["steps"])} for s in self.sections]

    def snapshot_between_delays(self):
        with self._lock:
            return list(self.delays_between)
