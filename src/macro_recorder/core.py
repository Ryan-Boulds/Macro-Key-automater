# macro_recorder/core.py
import time
import json
import threading
from typing import List, Dict, Any, Optional, Callable, Tuple

from .recorder import Recorder
from .merger import merge_steps
from .player import PlaybackExecutor


class MacroRecorderCore:
    def __init__(self):
        self.rows: List[List[Dict[str, Any]]] = []
        self.delays_within_rows: List[List[int]] = []
        self.delays_between_rows: List[int] = []
        self.recording = False
        self.active_section_index: Optional[Tuple[int, int]] = None

        self.ui_callback: Optional[Callable[[], None]] = None
        self.playback_ui_callback: Optional[Callable[[Tuple[int, int], int, bool], None]] = None

        self._lock = threading.Lock()
        self._last_ui_update = 0
        self._ui_update_interval = 0.1
        self.root = None  # Will be set by the UI (Tk root)

        # sub-components
        self._recorder = Recorder(self)
        self._player = PlaybackExecutor(self)

    # ------------------------------------------------------------------ #
    # UI notification (dynamic throttling)
    # ------------------------------------------------------------------ #
    def _notify_ui(self):
        now = time.time()
        interval = 0.5 if self.recording else 0.1   # 500 ms while recording
        if now - self._last_ui_update >= interval:
            if self.ui_callback:
                try:
                    self.ui_callback()
                    self._last_ui_update = now
                except Exception:
                    pass

    # ------------------------------------------------------------------ #
    # Gap management
    # ------------------------------------------------------------------ #
    def _ensure_gaps(self):
        """Make sure the gap-lists have the correct length after any structural change."""
        n_rows = len(self.rows)

        # between-row gaps
        self.delays_between_rows = (
            self.delays_between_rows[: max(0, n_rows - 1)]
            + [0] * max(0, (n_rows - 1) - len(self.delays_between_rows))
        )

        # within-row gaps
        while len(self.delays_within_rows) < n_rows:
            self.delays_within_rows.append([])

        for r in range(n_rows):
            n_secs = len(self.rows[r])
            self.delays_within_rows[r] = (
                self.delays_within_rows[r][: max(0, n_secs - 1)]
                + [0] * max(0, (n_secs - 1) - len(self.delays_within_rows[r]))
            )

    # ------------------------------------------------------------------ #
    # Section management (2-D)
    # ------------------------------------------------------------------ #
    def add_row(self):
        with self._lock:
            self.rows.append([])
            self.delays_within_rows.append([])
            self._ensure_gaps()
        self._notify_ui()

    def add_section(self, row_idx: int, name: str = "New Section") -> Tuple[int, int]:
        with self._lock:
            if 0 <= row_idx < len(self.rows):
                self.rows[row_idx].append({"name": name, "steps": []})
                self._ensure_gaps()
                return (row_idx, len(self.rows[row_idx]) - 1)
        self._notify_ui()
        return (-1, -1)

    def rename_section(self, idx: Tuple[int, int], name: str):
        with self._lock:
            row, col = idx
            if 0 <= row < len(self.rows) and 0 <= col < len(self.rows[row]):
                self.rows[row][col]["name"] = name
        self._notify_ui()

    def delete_section(self, idx: Tuple[int, int]):
        with self._lock:
            row, col = idx
            if 0 <= row < len(self.rows) and 0 <= col < len(self.rows[row]):
                del self.rows[row][col]
                if col < len(self.delays_within_rows[row]):
                    del self.delays_within_rows[row][col]
                if self.active_section_index == idx:
                    self.active_section_index = None
                self._ensure_gaps()
        self._notify_ui()

    def move_section_left(self, idx: Tuple[int, int]):
        with self._lock:
            row, col = idx
            if 0 <= row < len(self.rows) and 1 <= col < len(self.rows[row]):
                # swap sections
                self.rows[row][col - 1], self.rows[row][col] = (
                    self.rows[row][col],
                    self.rows[row][col - 1],
                )
                # swap within-row gaps (only if the gap exists)
                if (col - 1) < len(self.delays_within_rows[row]):
                    self.delays_within_rows[row][col - 1], self.delays_within_rows[row][col - 1] = (
                        self.delays_within_rows[row][col - 1],
                        self.delays_within_rows[row][col - 1],
                    )
                # adjust active index
                if self.active_section_index == idx:
                    self.active_section_index = (row, col - 1)
                elif self.active_section_index == (row, col - 1):
                    self.active_section_index = (row, col)
        self._notify_ui()

    def move_section_right(self, idx: Tuple[int, int]):
        with self._lock:
            row, col = idx
            if 0 <= row < len(self.rows) and 0 <= col < len(self.rows[row]) - 1:
                # swap sections
                self.rows[row][col + 1], self.rows[row][col] = (
                    self.rows[row][col],
                    self.rows[row][col + 1],
                )
                # swap within-row gaps
                if col < len(self.delays_within_rows[row]):
                    self.delays_within_rows[row][col], self.delays_within_rows[row][col] = (
                        self.delays_within_rows[row][col],
                        self.delays_within_rows[row][col],
                    )
                # adjust active index
                if self.active_section_index == idx:
                    self.active_section_index = (row, col + 1)
                elif self.active_section_index == (row, col + 1):
                    self.active_section_index = (row, col)
        self._notify_ui()

    # ------------------------------------------------------------------ #
    # Step editing
    # ------------------------------------------------------------------ #
    def add_delay_step(self, section_index: Tuple[int, int], delay_ms: int):
        row, col = section_index
        with self._lock:
            if 0 <= row < len(self.rows) and 0 <= col < len(self.rows[row]):
                self.rows[row][col]["steps"].append(
                    {"type": "delay", "delay": int(delay_ms), "unit": "ms"}
                )
        self._notify_ui()

    def add_typed_step(self, section_index: Tuple[int, int], chars: str, delay_ms: int):
        row, col = section_index
        with self._lock:
            if 0 <= row < len(self.rows) and 0 <= col < len(self.rows[row]):
                self.rows[row][col]["steps"].append(
                    {
                        "type": "typed",
                        "chars": chars,
                        "delays": [int(delay_ms)] * (len(chars) - 1),
                    }
                )
        self._notify_ui()

    def delete_step(self, section_index: Tuple[int, int], step_index: int):
        row, col = section_index
        with self._lock:
            if 0 <= row < len(self.rows) and 0 <= col < len(self.rows[row]):
                steps = self.rows[row][col]["steps"]
                if 0 <= step_index < len(steps):
                    del steps[step_index]
        self._notify_ui()

    def move_step_up(self, section_index: Tuple[int, int], step_index: int):
        row, col = section_index
        with self._lock:
            if 0 <= row < len(self.rows) and 0 <= col < len(self.rows[row]):
                steps = self.rows[row][col]["steps"]
                if 1 <= step_index < len(steps):
                    steps[step_index - 1], steps[step_index] = (
                        steps[step_index],
                        steps[step_index - 1],
                    )
        self._notify_ui()

    def move_step_down(self, section_index: Tuple[int, int], step_index: int):
        row, col = section_index
        with self._lock:
            if 0 <= row < len(self.rows) and 0 <= col < len(self.rows[row]):
                steps = self.rows[row][col]["steps"]
                if 0 <= step_index < len(steps) - 1:
                    steps[step_index + 1], steps[step_index] = (
                        steps[step_index],
                        steps[step_index + 1],
                    )
        self._notify_ui()

    def block_move_up(self, section_index: Tuple[int, int], start_idx: int, end_idx: int):
        row, col = section_index
        with self._lock:
            if 0 <= row < len(self.rows) and 0 <= col < len(self.rows[row]):
                steps = self.rows[row][col]["steps"]
                if 0 <= start_idx <= end_idx < len(steps) and start_idx > 0:
                    block = steps[start_idx : end_idx + 1]
                    del steps[start_idx : end_idx + 1]
                    steps[start_idx - 1 : start_idx - 1] = block
        self._notify_ui()

    def block_move_down(self, section_index: Tuple[int, int], start_idx: int, end_idx: int):
        row, col = section_index
        with self._lock:
            if 0 <= row < len(self.rows) and 0 <= col < len(self.rows[row]):
                steps = self.rows[row][col]["steps"]
                if 0 <= start_idx <= end_idx < len(steps) - 1:
                    block = steps[start_idx : end_idx + 1]
                    del steps[start_idx : end_idx + 1]
                    steps[end_idx + 1 : end_idx + 1] = block
        self._notify_ui()

    def edit_delay(self, section_index: Tuple[int, int], step_index: int, new_delay_ms: int):
        row, col = section_index
        with self._lock:
            if 0 <= row < len(self.rows) and 0 <= col < len(self.rows[row]):
                steps = self.rows[row][col]["steps"]
                if 0 <= step_index < len(steps):
                    step = steps[step_index]
                    if step.get("type") == "delay":
                        step["delay"] = int(new_delay_ms)
                        step["unit"] = "ms"
        self._notify_ui()

    def set_gap_delay(self, gap_index: Tuple[int, int], ms: int):
        row, gap = gap_index
        with self._lock:
            if 0 <= row < len(self.delays_within_rows) and 0 <= gap < len(self.delays_within_rows[row]):
                self.delays_within_rows[row][gap] = int(ms)
        self._notify_ui()

    def set_between_row_delay(self, row_gap_index: int, ms: int):
        with self._lock:
            if 0 <= row_gap_index < len(self.delays_between_rows):
                self.delays_between_rows[row_gap_index] = int(ms)
        self._notify_ui()

    def clear_all(self):
        with self._lock:
            self.rows.clear()
            self.delays_within_rows.clear()
            self.delays_between_rows.clear()
            self.active_section_index = None
        self._notify_ui()

    # ------------------------------------------------------------------ #
    # Recording delegation
    # ------------------------------------------------------------------ #
    def start_recording(self, section_index: Tuple[int, int]):
        row, col = section_index
        if not (0 <= row < len(self.rows) and 0 <= col < len(self.rows[row])):
            return
        self._recorder.start(section_index)
        self._notify_ui()

    def stop_recording(self):
        self._recorder.stop()
        self._notify_ui()

    # ------------------------------------------------------------------ #
    # Playback delegation
    # ------------------------------------------------------------------ #
    def play_all(self, stop_event: Optional[threading.Event] = None):
        self._player.play_all(stop_event)

    # ------------------------------------------------------------------ #
    # Persistence
    # ------------------------------------------------------------------ #
    def save_macro(self, path: str):
        with self._lock:
            data = {
                "rows": self.rows,
                "delays_within_rows": self.delays_within_rows,
                "delays_between_rows": self.delays_between_rows,
            }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def load_macro(self, path: str):
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        with self._lock:
            self.rows = data.get("rows", [])
            self.delays_within_rows = data.get(
                "delays_within_rows", [[] for _ in self.rows]
            )
            self.delays_between_rows = data.get(
                "delays_between_rows", [0] * max(0, len(self.rows) - 1)
            )
            self._ensure_gaps()
        self._notify_ui()

    # ------------------------------------------------------------------ #
    # Snapshots (used by playback & UI)
    # ------------------------------------------------------------------ #
    def snapshot_rows(self):
        with self._lock:
            return [
                [{"name": s["name"], "steps": [step.copy() for step in s["steps"]]} for s in row]
                for row in self.rows
            ]

    def snapshot_within_rows(self):
        with self._lock:
            return [list(gaps) for gaps in self.delays_within_rows]

    def snapshot_between_rows(self):
        with self._lock:
            return list(self.delays_between_rows)