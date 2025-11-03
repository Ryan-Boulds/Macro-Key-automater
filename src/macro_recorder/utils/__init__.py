# macro_recorder/utils/__init__.py
# Keep your existing `normalize_key` implementation here
# Example (replace with your real version):
from pynput.keyboard import Key

def normalize_key(key):
    """Convert pynput Key / KeyCode to a simple string."""
    if hasattr(key, "char") and key.char:
        return key.char.lower()
    mapping = {
        Key.space: "space",
        Key.enter: "enter",
        Key.tab: "tab",
        Key.backspace: "backspace",
        Key.delete: "delete",
        Key.shift: "shift",
        Key.shift_r: "shift",
        Key.ctrl: "ctrl",
        Key.ctrl_r: "ctrl",
        Key.alt: "alt",
        Key.alt_r: "alt",
        Key.cmd: "cmd",
        Key.cmd_r: "cmd_r",
        Key.menu: "cmd",
        Key.left: "left",
        Key.right: "right",
        Key.up: "up",
        Key.down: "down",
        Key.esc: "esc",
    }
    return mapping.get(key, str(key).replace("Key.", ""))