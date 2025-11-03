# src/section_manager.py
class SectionManager:
    def __init__(self, app):
        self.app = app

    def add(self, row_idx=None, name=None):
        """Add a new section (column) to the given row."""
        if row_idx is None:
            row_idx = self.app.active_row_index
        if row_idx is None:
            row_idx = 0
            self.app.add_row()
        if name is None:
            name = f"Section {len(self.app.recorder.rows[row_idx]) + 1}"
        idx = self.app.recorder.add_section(row_idx, name)
        self.app.active_section_index = idx
        self.app.render_sections()
        return idx

    def delete(self, idx):
        row_idx, col_idx = idx
        if not (0 <= row_idx < len(self.app.recorder.rows) and
                0 <= col_idx < len(self.app.recorder.rows[row_idx])):
            return
        self.app.recorder.delete_section(idx)
        if self.app.active_section_index == idx:
            self.app.active_section_index = None
        # clear last-recorded step if it belonged to the deleted section
        if self.app.last_recorded_step and self.app.last_recorded_step[:2] == idx:
            self.app.last_recorded_step = None
        # remove any selected steps that were in the deleted section
        self.app.selection.selected_indices = {
            k for k in self.app.selection.selected_indices if k[0] != idx
        }
        self.app.render_sections()

    def select(self, idx):
        row_idx, _ = idx
        self.app.active_row_index = row_idx
        self.app.active_section_index = idx
        self.app.render_sections()

    def move_left(self, idx):
        _, col_idx = idx
        if col_idx > 0:
            self.app.recorder.move_section_left(idx)
            self.app.render_sections()

    def move_right(self, idx):
        row_idx, col_idx = idx
        if col_idx < len(self.app.recorder.rows[row_idx]) - 1:
            self.app.recorder.move_section_right(idx)
            self.app.render_sections()

    def rename(self, idx, name):
        self.app.recorder.rename_section(idx, name)
        self.app.render_sections()