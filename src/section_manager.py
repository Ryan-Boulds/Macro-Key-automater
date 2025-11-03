# src/section_manager.py
import tkinter.messagebox as messagebox


class SectionManager:
    def __init__(self, app):
        self.app = app

    def add(self, name=None):
        if name is None:
            name = f"Section {len(self.app.recorder.sections) + 1}"
        idx = self.app.recorder.add_section(name)
        self.app.active_section_index = idx
        self.app.ui.refresh()  # ← Fixed
        return idx

    def delete(self, idx):
        if not (0 <= idx < len(self.app.recorder.sections)):
            return
        self.app.recorder.delete_section(idx)
        if self.app.active_section_index is not None:
            if self.app.active_section_index >= len(self.app.recorder.sections):
                self.app.active_section_index = max(0, len(self.app.recorder.sections) - 1)
        if self.app.last_recorded_step and self.app.last_recorded_step[0] == idx:
            self.app.last_recorded_step = None
        self.app.selection.selected_steps = {
            k: v for k, v in self.app.selection.selected_steps.items() if k[0] != idx
        }
        self.app.ui.refresh()  # ← Fixed

    def select(self, idx):
        self.app.active_section_index = idx
        self.app.ui.refresh()  # ← Fixed

    def move_left(self, idx):
        self.app.recorder.move_section_left(idx)
        self.app.ui.refresh()  # ← Fixed

    def move_right(self, idx):
        self.app.recorder.move_section_right(idx)
        self.app.ui.refresh()  # ← Fixed

    def rename(self, idx, name):
        self.app.recorder.rename_section(idx, name)