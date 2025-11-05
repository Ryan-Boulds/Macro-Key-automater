# macro_recorder/merger.py
from typing import List, Dict, Any


def merge_steps(steps: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Convert raw press/release/delay/mouse_move into high-level actions:
      - mouse_click
      - mouse_drag
      - typed
      - key_group
      - delay
    """
    new_steps: List[Dict[str, Any]] = []
    i = 0

    # Keys that **break** a typed string
    NON_CHAR_KEYS = {
        "enter", "tab", "esc", "backspace", "delete",
        "up", "down", "left", "right",
        "page_up", "page_down", "home", "end", "insert",
        "f1", "f2", "f3", "f4", "f5", "f6",
        "f7", "f8", "f9", "f10", "f11", "f12",
    }

    # Pure modifiers (for key_group)
    MODIFIER_KEYS = {"cmd", "cmd_r", "win", "ctrl", "alt", "shift"}

    shift_map = {
        "1": "!", "2": "@", "3": "#", "4": "$", "5": "%",
        "6": "^", "7": "&", "8": "*", "9": "(", "0": ")",
        "-": "_", "=": "+", ".": "."
    }
    allowed_chars = set("abcdefghijklmnopqrstuvwxyz0123456789. ")

    def _flush_typed(chars: List[str], delays: List[int]):
        if chars:
            new_steps.append({
                "type": "typed",
                "chars": "".join(chars),
                "delays": delays or [15] * (len(chars) - 1)
            })

    while i < len(steps):
        step = steps[i]

        # ------------------------------------------------------------------
        # Mouse Press → start of click or drag
        # ------------------------------------------------------------------
        if step["type"] == "mouse_press":
            drag_step = {
                "type": "mouse_drag",
                "button": step["button"],
                "start_x": step["x"],
                "start_y": step["y"],
                "path": [(step["x"], step["y"])],
                "duration_ms": 0,
            }
            i += 1
            released = False

            while i < len(steps):
                nxt = steps[i]

                if nxt["type"] == "delay":
                    drag_step["duration_ms"] += nxt["delay"]
                    i += 1
                elif nxt["type"] == "mouse_move":
                    x, y = nxt["x"], nxt["y"]
                    if (x, y) != drag_step["path"][-1]:  # avoid duplicates
                        drag_step["path"].append((x, y))
                    i += 1
                elif nxt["type"] == "mouse_release" and nxt["button"] == step["button"]:
                    drag_step["end_x"] = nxt["x"]
                    drag_step["end_y"] = nxt["y"]
                    released = True
                    i += 1
                    break
                else:
                    break

            # Finalize drag
            if released and len(drag_step["path"]) > 1:
                # Downsample path for smooth playback (max 50 points)
                if len(drag_step["path"]) > 50:
                    step = len(drag_step["path"]) // 50
                    drag_step["path"] = drag_step["path"][::step]
                new_steps.append(drag_step)
            else:
                # No movement or no release → treat as click
                new_steps.append({
                    "type": "mouse_click",
                    "button": step["button"],
                    "x": step["x"],
                    "y": step["y"],
                    "hold_ms": drag_step["duration_ms"],
                })
            continue

        # ------------------------------------------------------------------
        # Other non-keyboard steps
        # ------------------------------------------------------------------
        if step["type"] not in ("press", "release", "delay", "mouse_move"):
            new_steps.append(step)
            i += 1
            continue

        # ------------------------------------------------------------------
        # Keyboard block
        # ------------------------------------------------------------------
        sub_steps = []
        typed_chars, typed_delays = [], []
        pressed_keys = set()
        shift_pressed = False
        j = i

        while j < len(steps) and steps[j]["type"] in ("press", "release", "delay", "mouse_move"):
            if steps[j]["type"] == "mouse_move":
                break  # mouse move breaks keyboard block
            sub_steps.append(steps[j])
            s = steps[j]
            if s["type"] == "press":
                pressed_keys.add(s["key"])
                if s["key"] == "shift":
                    shift_pressed = True
            elif s["type"] == "release" and s["key"] in pressed_keys:
                pressed_keys.remove(s["key"])
                if s["key"] == "shift":
                    shift_pressed = False
            j += 1

        k = 0
        current_key_group = []

        while k < len(sub_steps):
            s = sub_steps[k]

            # Delay
            if s["type"] == "delay":
                current_key_group.append(s)
                k += 1
                continue

            # Modifier press
            if s["type"] == "press" and s["key"] in MODIFIER_KEYS:
                current_key_group.append(s)
                k += 1
                continue

            # Modifier release
            if s["type"] == "release" and s["key"] in MODIFIER_KEYS:
                current_key_group.append(s)
                pressed_keys.discard(s["key"])
                k += 1
                if not any(p in MODIFIER_KEYS for p in pressed_keys):
                    if current_key_group:
                        new_steps.append({"type": "key_group", "sub_steps": current_key_group})
                    current_key_group = []
                continue

            # Non-char key → split typed
            if s["type"] == "press" and s["key"] in NON_CHAR_KEYS:
                _flush_typed(typed_chars, typed_delays)
                typed_chars.clear()
                typed_delays.clear()
                current_key_group.append(s)
                k += 1
                continue

            # Printable char
            if s["type"] == "press" and (
                s["key"] in allowed_chars or
                s["key"] == "space" or
                (shift_pressed and s["key"] in shift_map)
            ):
                if current_key_group:
                    new_steps.append({"type": "key_group", "sub_steps": current_key_group})
                    current_key_group = []

                char = shift_map.get(s["key"], s["key"]) if shift_pressed and s["key"] in shift_map else (
                    " " if s["key"] == "space" else s["key"]
                )
                typed_chars.append(char)
                k += 1

                # Consume release + delay
                while k < len(sub_steps):
                    nxt = sub_steps[k]
                    if nxt["type"] == "release" and nxt["key"] in (s["key"], "shift", "space"):
                        if nxt["key"] == "shift":
                            shift_pressed = False
                        k += 1
                    elif nxt["type"] == "delay":
                        typed_delays.append(nxt["delay"])
                        k += 1
                        break
                    else:
                        break
                continue

            k += 1

        # Final flush
        _flush_typed(typed_chars, typed_delays)
        if current_key_group:
            new_steps.append({"type": "key_group", "sub_steps": current_key_group})

        i = j

    return new_steps