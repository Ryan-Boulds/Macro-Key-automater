# src/selection_manager.py
import tkinter.messagebox as messagebox


class SelectionManager:
    def __init__(self, app):
        self.app = app
        self.selected_steps = {}
        self.last_clicked = None

    def clear(self):
        for (si, sti), lbl in self.selected_steps.items():
            bg = "#FFFF99" if (si, sti) == self.app.last_recorded_step else "white"
            lbl.config(bg=bg)
        self.selected_steps.clear()

    def toggle(self, si, sti, lbl, ctrl_held):
        key = (si, sti)
        if ctrl_held:
            if key in self.selected_steps:
                bg = "#FFFF99" if key == self.app.last_recorded_step else "white"
                self.selected_steps[key].config(bg=bg)
                del self.selected_steps[key]
            else:
                self.selected_steps[key] = lbl
                lbl.config(bg="#D3D3D3")
        else:
            self.clear()
            self.selected_steps[key] = lbl
            lbl.config(bg="#D3D3D3")
        self.last_clicked = key

    def move_up(self):
        self._move_steps("Up")

    def move_down(self):
        self._move_steps("Down")

    def _move_steps(self, direction):
        if not self.selected_steps:
            if self.last_clicked:
                si, sti = self.last_clicked
                getattr(self.app.recorder, f"move_step_{direction.lower()}")(si, sti)
            return

        sections = {}
        for (si, sti) in self.selected_steps:
            sections.setdefault(si, []).append(sti)

        new_selected = {}
        for si, indices in sections.items():
            indices = sorted(indices)
            if len(indices) == max(indices) - min(indices) + 1:  # consecutive
                start, end = indices[0], indices[-1]
                if direction == "Up" and start > 0:
                    self.app.recorder.block_move_up(si, start, end)
                    for idx in indices:
                        new_selected[(si, idx - 1)] = self.selected_steps[(si, idx)]
                elif direction == "Down" and end < len(self.app.recorder.sections[si]["steps"]) - 1:
                    self.app.recorder.block_move_down(si, start, end)
                    for idx in indices:
                        new_selected[(si, idx + 1)] = self.selected_steps[(si, idx)]
                else:
                    new_selected.update({(si, i): self.selected_steps[(si, i)] for i in indices})
            else:
                for idx in sorted(indices, reverse=(direction == "Up")):
                    if (direction == "Up" and idx > 0) or (direction == "Down" and idx < len(self.app.recorder.sections[si]["steps"]) - 1):
                        getattr(self.app.recorder, f"move_step_{direction.lower()}")(si, idx)
                        new_idx = idx - 1 if direction == "Up" else idx + 1
                        new_selected[(si, new_idx)] = self.selected_steps[(si, idx)]
                    else:
                        new_selected[(si, idx)] = self.selected_steps[(si, idx)]

        self.selected_steps = new_selected
        self.app.ui.refresh()  # ← Fixed

    def delete_selected(self):
        selected = sorted(self.selected_steps.keys(), key=lambda x: (x[0], x[1]), reverse=True)
        for si, sti in selected:
            self.app.recorder.delete_step(si, sti)
            if (si, sti) == self.app.last_recorded_step:
                self.app.last_recorded_step = None
        self.selected_steps.clear()
        self.app.ui.refresh()  # ← Fixed