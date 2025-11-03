# src/ui_updater.py


class UIUpdater:
    def __init__(self, app):
        self.app = app
        self.pending_update = False

    def refresh(self):
        if self.app._is_visible():
            self.app.root.after(0, self.app.render_sections)
        else:
            self.pending_update = True

    def on_focus(self):
        if self.pending_update:
            self.app.render_sections()
            self.pending_update = False

    def scroll_to(self, widget):
        x = y = 0
        w = widget
        while w != self.app.sections_frame:
            x += w.winfo_x()
            y += w.winfo_y()
            w = w.master
        x += widget.winfo_width() / 2
        y += widget.winfo_height() / 2

        cw, ch = self.app.canvas.winfo_width(), self.app.canvas.winfo_height()
        fw, fh = self.app.sections_frame.winfo_width(), self.app.sections_frame.winfo_height()

        frac_x = max(0, min(1, (x - cw / 2) / fw))
        frac_y = max(0, min(1, (y - ch / 2) / fh))

        self.app.canvas.xview_moveto(frac_x)
        self.app.canvas.yview_moveto(frac_y)

    def highlight_playback(self, sec_idx, step_idx, active):
        def do():
            bg = "#ADD8E6" if active else "white"
            widget = None
            if step_idx >= 0 and 0 <= sec_idx < len(self.app.step_labels):
                if 0 <= step_idx < len(self.app.step_labels[sec_idx]):
                    lbl = self.app.step_labels[sec_idx][step_idx]
                    if (sec_idx, step_idx) == self.app.last_recorded_step:
                        bg = "#FFFF99" if not active else "#ADD8E6"
                    elif (sec_idx, step_idx) in self.app.selection.selected_steps:
                        bg = "#D3D3D3" if not active else "#ADD8E6"
                    lbl.config(bg=bg)
                    widget = lbl
            else:
                if 0 <= sec_idx < len(self.app.gap_chips):
                    self.app.gap_chips[sec_idx].config(bg=bg)
                    widget = self.app.gap_chips[sec_idx]
            if active and widget:
                self.scroll_to(widget)
        self.app.root.after(0, do)