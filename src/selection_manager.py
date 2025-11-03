# src/selection_manager.py
import tkinter.messagebox as messagebox


class SelectionManager:
    def __init__(self, app):
        self.app = app
        self.selected_indices = set()   # set of ((row, sec), step_idx)
        self.last_clicked = None

    def clear(self):
        for key in list(self.selected_indices):
            (si, sti) = key
            try:
                lbl = self.app.step_labels[si[0]][si[1]][sti]
                bg = "#FFFF99" if key == self.app.last_recorded_step else "white"
                lbl.config(bg=bg)
            except Exception:
                pass
        self.selected_indices.clear()

    def toggle(self, si, sti, lbl, ctrl_held):
        key = (si, sti)
        if ctrl_held:
            if key in self.selected_indices:
                bg = "#FFFF99" if key == self.app.last_recorded_step else "white"
                lbl.config(bg=bg)
                self.selected_indices.remove(key)
            else:
                self.selected_indices.add(key)
                lbl.config(bg="#D3D3D3")
        else:
            self.clear()
            self.selected_indices.add(key)
            lbl.config(bg="#D3D3D3")
        self.last_clicked = key

    def move_up(self):
        self._move_steps("Up")

    def move_down(self):
        self._move_steps("Down")

    def _move_steps(self, direction):
        if not self.selected_indices:
            if self.last_clicked:
                si, sti = self.last_clicked
                getattr(self.app.recorder, f"move_step_{direction.lower()}")(si, sti)
            return

        sections = {}
        for (si, sti) in self.selected_indices:
            sections.setdefault(si, []).append(sti)

        new_set = set()
        for si, indices in sections.items():
            indices = sorted(indices)
            if len(indices) == max(indices) - min(indices) + 1:   # consecutive block
                start, end = indices[0], indices[-1]
                if direction == "Up" and start > 0:
                    self.app.recorder.block_move_up(si, start, end)
                    for i in indices:
                        new_set.add((si, i - 1))
                elif direction == "Down" and end < len(self.app.recorder.rows[si[0]][si[1]]["steps"]) - 1:
                    self.app.recorder.block_move_down(si, start, end)
                    for i in indices:
                        new_set.add((si, i + 1))
                else:
                    for i in indices:
                        new_set.add((si, i))
            else:
                for idx in sorted(indices, reverse=(direction == "Up")):
                    if (direction == "Up" and idx > 0) or \
                       (direction == "Down" and idx < len(self.app.recorder.rows[si[0]][si[1]]["steps"]) - 1):
                        getattr(self.app.recorder, f"move_step_{direction.lower()}")(si, idx)
                        new_set.add((si, idx - 1 if direction == "Up" else idx + 1))
                    else:
                        new_set.add((si, idx))

        self.selected_indices = new_set
        self.app.render_sections()

    def delete_selected(self):
        selected = sorted(self.selected_indices, key=lambda x: (x[0], x[1]), reverse=True)
        for si, sti in selected:
            self.app.recorder.delete_step(si, sti)
            if (si, sti) == self.app.last_recorded_step:
                self.app.last_recorded_step = None
        self.selected_indices.clear()
        self.app.render_sections()