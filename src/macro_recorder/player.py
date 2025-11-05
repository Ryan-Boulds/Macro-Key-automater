# macro_recorder/player.py
import time
import threading
import ctypes
from ctypes import wintypes, c_ulong, POINTER, sizeof, byref

import pyautogui
from typing import Optional, Dict, Any

# ----------------------------------------------------------------------
# WinAPI helpers – emergency clean-up only
# ----------------------------------------------------------------------
user32 = ctypes.windll.user32
GetKeyState = user32.GetKeyState
GetKeyState.restype = wintypes.SHORT

VK_LWIN = 0x5B
VK_RWIN = 0x5C

def _is_win_down() -> bool:
    return bool(GetKeyState(VK_LWIN) & 0x8000) or bool(GetKeyState(VK_RWIN) & 0x8000)

def _force_win_release():
    extra = c_ulong(0)
    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD),
                    ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD),
                    ("dwExtraInfo", POINTER(c_ulong))]
    class INPUT(ctypes.Structure):
        _fields_ = [("type", wintypes.DWORD), ("ki", KEYBDINPUT)]
    def send_up(vk):
        ki = KEYBDINPUT(vk, 0, 0x0002, 0, byref(extra))
        inp = INPUT(0, ki)
        user32.SendInput(1, byref(inp), sizeof(inp))
    send_up(VK_LWIN)
    send_up(VK_RWIN)
    time.sleep(0.02)

# ----------------------------------------------------------------------
# PlaybackExecutor – 2-D ROWS version
# ----------------------------------------------------------------------
class PlaybackExecutor:
    def __init__(self, core):
        self.core = core

    def play_all(self, stop_event: Optional[threading.Event] = None):
        # Use 2-D API
        snapshot_rows = self.core.snapshot_rows()
        gaps_within = self.core.snapshot_within_rows()
        gaps_between = self.core.snapshot_between_rows()

        try:
            for row_idx, row in enumerate(snapshot_rows):
                for sec_idx, section in enumerate(row):
                    if stop_event and stop_event.is_set():
                        return
                    for step_idx, action in enumerate(section["steps"]):
                        if stop_event and stop_event.is_set():
                            return
                        self.core._playback_notify((row_idx, sec_idx), step_idx, True)
                        self._execute_action(action, stop_event)
                        self.core._playback_notify((row_idx, sec_idx), step_idx, False)

                    # Gap inside row
                    if sec_idx < len(row) - 1:
                        delay_ms = (gaps_within[row_idx][sec_idx]
                                    if row_idx < len(gaps_within) and sec_idx < len(gaps_within[row_idx])
                                    else 0)
                        if delay_ms > 0:
                            self.core._playback_notify((row_idx, sec_idx), -1, True)
                            self._sleep_with_interrupt(delay_ms / 1000.0, stop_event)
                            self.core._playback_notify((row_idx, sec_idx), -1, False)

                # Gap between rows
                if row_idx < len(snapshot_rows) - 1:
                    delay_ms = gaps_between[row_idx] if row_idx < len(gaps_between) else 0
                    if delay_ms > 0:
                        self.core._playback_notify((row_idx, -1), -1, True)
                        self._sleep_with_interrupt(delay_ms / 1000.0, stop_event)
                        self.core._playback_notify((row_idx, -1), -1, False)
        finally:
            if _is_win_down():
                pyautogui.keyUp("winleft")
                time.sleep(0.05)
                if _is_win_down():
                    _force_win_release()

    @staticmethod
    def _sleep_with_interrupt(seconds: float, stop_event: Optional[threading.Event] = None):
        start = time.time()
        while time.time() - start < seconds:
            if stop_event and stop_event.is_set():
                return
            time.sleep(0.01)

    def _execute_action(self, action: Dict[str, Any], stop_event: Optional[threading.Event]):
        t = action.get("type")

        if t == "delay":
            unit = action.get("unit", "ms")
            secs = (
                action["delay"] / 1000 if unit == "ms" else
                action["delay"] if unit == "secs" else
                action["delay"] * 60 if unit == "mins" else
                action["delay"] * 3600
            )
            self._sleep_with_interrupt(secs, stop_event)

        elif t == "press" and action.get("key") in ("win", "cmd", "cmd_r"):
            pyautogui.press('win')           # Atomic: opens Start, never stuck
            time.sleep(0.12)                 # Let Start menu appear

        elif t == "press":
            pyautogui.keyDown(action.get("key"))

        elif t == "release":
            pyautogui.keyUp(action.get("key"))

        elif t == "mouse_press":
            pyautogui.moveTo(action["x"], action["y"])
            pyautogui.mouseDown(button=action["button"])

        elif t == "mouse_release":
            pyautogui.moveTo(action["x"], action["y"])
            pyautogui.mouseUp(button=action["button"])

        elif t == "key_group":
            for sub in action.get("sub_steps", []):
                self._execute_action(sub, stop_event)

        elif t == "mouse_click":
            pyautogui.moveTo(action["x"], action["y"])
            pyautogui.mouseDown(button=action["button"])
            self._sleep_with_interrupt(action.get("hold_ms", 0) / 1000.0, stop_event)
            rx = action.get("release_x", action["x"])
            ry = action.get("release_y", action["y"])
            pyautogui.moveTo(rx, ry)
            pyautogui.mouseUp(button=action["button"])

        elif t == "typed":
            chars = action.get("chars", "")
            delays = action.get("delays", [])
            for i, ch in enumerate(chars):
                if ch in ("%","^","+","(",")","{","}","~"):
                    pyautogui.hotkey("shift", ch)
                elif ch == " ":
                    pyautogui.press("space")
                else:
                    pyautogui.press(ch)
                if i < len(chars) - 1:
                    d = delays[i] if i < len(delays) else 15
                    self._sleep_with_interrupt(d / 1000.0, stop_event)
        elif t == "mouse_drag":
            pyautogui.moveTo(action["start_x"], action["start_y"])
            pyautogui.mouseDown(button=action["button"])
            duration_sec = action["duration_ms"] / 1000.0
            if duration_sec > 0:
                # Smooth move along path or direct
                if "path" in action and len(action["path"]) > 1:
                    for px, py in action["path"][1:]:
                        pyautogui.moveTo(px, py, duration=0)
                else:
                    pyautogui.moveTo(action["end_x"], action["end_y"], duration=duration_sec)
            else:
                pyautogui.moveTo(action["end_x"], action["end_y"])
            pyautogui.mouseUp(button=action["button"])