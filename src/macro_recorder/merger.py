# macro_recorder/merger.py
from typing import List, Dict, Any


def merge_steps(steps: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Takes a raw list of press/release/delay steps and returns a compacted list
    containing:
      * mouse_click (with optional hold_ms + release coordinates)
      * typed (merged characters + per-char delays)
      * key_group (modifier combos)
      * plain delay steps
    """
    new_steps: List[Dict[str, Any]] = []
    i = 0
    modifier_keys = {
        "cmd", "cmd_r", "win", "ctrl", "alt", "shift",
        "enter", "tab", "esc", "backspace", "delete"
    }
    shift_map = {
        "1": "!", "2": "@", "3": "#", "4": "$", "5": "%",
        "6": "^", "7": "&", "8": "*", "9": "(", "0": ")",
        "-": "_", "=": "+", ".": "."
    }
    allowed_chars = set("abcdefghijklmnopqrstuvwxyz0123456789. ")

    while i < len(steps):
        step = steps[i]

        # ------------------------------------------------------------------
        # Non-keyboard actions (mouse, delay, etc.)
        # ------------------------------------------------------------------
        if step["type"] not in ("press", "release", "delay"):
            if step["type"] == "mouse_press":
                mouse_step = {
                    "type": "mouse_click",
                    "button": step["button"],
                    "x": step["x"],
                    "y": step["y"],
                    "hold_ms": 0,
                }
                i += 1
                while i < len(steps):
                    nxt = steps[i]
                    if nxt["type"] == "delay":
                        mouse_step["hold_ms"] += nxt["delay"]
                        i += 1
                    elif (nxt["type"] == "mouse_release" and
                          nxt["button"] == step["button"]):
                        mouse_step["release_x"] = nxt["x"]
                        mouse_step["release_y"] = nxt["y"]
                        i += 1
                        break
                    else:
                        break
                new_steps.append(mouse_step)
            else:
                new_steps.append(step)
                i += 1
            continue

        # ------------------------------------------------------------------
        # Keyboard sequence handling
        # ------------------------------------------------------------------
        sub_steps: List[Dict[str, Any]] = []
        typed_chars: List[str] = []
        typed_delays: List[int] = []
        pressed_keys: set = set()
        shift_pressed = False
        j = i
        while j < len(steps):
            s = steps[j]
            if s["type"] in ("press", "release", "delay"):
                sub_steps.append(s)
                if s["type"] == "press":
                    pressed_keys.add(s["key"])
                    if s["key"] == "shift":
                        shift_pressed = True
                elif s["type"] == "release" and s["key"] in pressed_keys:
                    pressed_keys.remove(s["key"])
                    if s["key"] == "shift":
                        shift_pressed = False
                j += 1
            else:
                break

        # ---- process the collected sub_steps --------------------------------
        k = 0
        current_key_group: List[Dict[str, Any]] = []
        last_delay = 0

        while k < len(sub_steps):
            s = sub_steps[k]

            if s["type"] == "delay":
                last_delay = s["delay"]
                current_key_group.append(s)
                k += 1
                continue

            if s["type"] == "press" and s["key"] in modifier_keys:
                current_key_group.append(s)
                k += 1
                continue

            if s["type"] == "release" and s["key"] in modifier_keys and s["key"] in pressed_keys:
                current_key_group.append(s)
                pressed_keys.remove(s["key"])
                k += 1
                # flush key_group when all modifiers are released
                if not pressed_keys or all(p in modifier_keys for p in pressed_keys):
                    if current_key_group and not all(
                        ss.get("key") == "shift" or ss["type"] == "delay"
                        for ss in current_key_group
                    ):
                        new_steps.append({"type": "key_group", "sub_steps": current_key_group})
                    current_key_group = []
                continue

            # ---- regular printable character ---------------------------------
            if s["type"] == "press" and (
                s["key"] in allowed_chars or
                s["key"] == "space" or
                (shift_pressed and s["key"] in shift_map)
            ):
                # flush any pending key_group first
                if current_key_group and not all(
                    ss.get("key") == "shift" or ss["type"] == "delay"
                    for ss in current_key_group
                ):
                    new_steps.append({"type": "key_group", "sub_steps": current_key_group})
                    current_key_group = []

                char = shift_map[s["key"]] if shift_pressed and s["key"] in shift_map else (
                    " " if s["key"] == "space" else s["key"]
                )
                typed_chars.append(char)
                k += 1

                # consume release + optional delay
                while k < len(sub_steps):
                    nxt = sub_steps[k]
                    if nxt["type"] == "release" and (
                        nxt["key"] == s["key"] or
                        (nxt["key"] == "shift" and shift_pressed) or
                        nxt["key"] == "space"
                    ):
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

            k += 1  # unknown step – skip

        # ---- final typed block ---------------------------------------------
        if typed_chars:
            new_steps.append({
                "type": "typed",
                "chars": "".join(typed_chars),
                "delays": typed_delays or [15] * (len(typed_chars) - 1)
            })

        # ---- leftover key_group --------------------------------------------
        if current_key_group and not all(
            ss.get("key") == "shift" or ss["type"] == "delay"
            for ss in current_key_group
        ):
            new_steps.append({"type": "key_group", "sub_steps": current_key_group})

        i = j

    return new_steps