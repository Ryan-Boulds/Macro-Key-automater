# src/ui_updater.py
import tkinter as tk


class UIUpdater:
    def __init__(self, app):
        self.app = app
        self.after_id = None

    def on_focus(self):
        self.refresh()

    def refresh(self):
        if self.after_id:
            self.app.root.after_cancel(self.after_id)
        interval = 100 if not self.app.recorder.recording else 500  # Longer during recording
        self.after_id = self.app.root.after(interval, self.do)  # Updated: Dynamic delay

    def do(self):
        self.after_id = None

        # playback highlight
        if hasattr(self.app, "playback") and self.app.playback.current_step:
            r, s, st = self.app.playback.current_step
            if st >= 0 and r < len(self.app.step_labels) and \
               s < len(self.app.step_labels[r]) and \
               st < len(self.app.step_labels[r][s]):
                self.app.step_labels[r][s][st].config(bg="#90EE90")   # light green

        # last-recorded highlight
        if self.app.last_recorded_step:
            r, s, st = self.app.last_recorded_step
            if r < len(self.app.step_labels) and \
               s < len(self.app.step_labels[r]) and \
               st < len(self.app.step_labels[r][s]):
                self.app.step_labels[r][s][st].config(bg="#FFFF99")   # yellow

        # selected steps
        for (si, sti) in self.app.selection.selected_indices:
            r, s = si
            if r < len(self.app.step_labels) and \
               s < len(self.app.step_labels[r]) and \
               sti < len(self.app.step_labels[r][s]):
                self.app.step_labels[r][s][sti].config(bg="#D3D3D3")

        # reset everything else
        for r_idx, row in enumerate(self.app.step_labels):
            for s_idx, sec in enumerate(row):
                for st_idx, lbl in enumerate(sec):
                    if (r_idx, s_idx, st_idx) != self.app.last_recorded_step and \
                       ((r_idx, s_idx), st_idx) not in self.app.selection.selected_indices and \
                       (hasattr(self.app.playback, "current_step") and
                        self.app.playback.current_step != (r_idx, s_idx, st_idx)):
                        lbl.config(bg="white")