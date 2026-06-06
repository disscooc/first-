# gui/autocomplete.py -- non-blocking autocomplete input
import tkinter as tk
from tkinter import ttk


class AutocompleteInput(ttk.Frame):
    def __init__(
        self,
        parent,
        textvariable,
        get_suggestions,
        get_all_suggestions=None,
        on_select=None,
        on_select_item=None,
        width=20,
        max_items=8,
        style=None,
        show_button=False,
        min_chars=1,
    ):
        frame_options = {"style": style} if style else {}
        super().__init__(parent, **frame_options)
        self.variable = textvariable
        self.get_suggestions = get_suggestions
        self.get_all_suggestions = get_all_suggestions or (lambda: [])
        self.on_select = on_select
        self.on_select_item = on_select_item
        self.max_items = max_items
        self.min_chars = min_chars
        self.items = []
        self.options = []
        self._trace_busy = False
        self._hide_after_id = None

        self.columnconfigure(0, weight=1)
        self.entry = ttk.Entry(self, textvariable=self.variable, width=width)
        self.entry.grid(row=0, column=0, sticky="ew")
        self.button = None
        if show_button:
            self.button = ttk.Button(self, text="⌄", width=2, command=self.toggle_all)
            self.button.grid(row=0, column=1, sticky="ns", padx=(1, 0))

        self.popup_parent = self.winfo_toplevel()
        self.popup = tk.Frame(self.popup_parent, bg="#D9E2EC", borderwidth=1, relief="solid")
        self.listbox = tk.Listbox(
            self.popup,
            activestyle="none",
            exportselection=False,
            borderwidth=0,
            highlightthickness=0,
            background="#FFFFFF",
            foreground="#202124",
            selectbackground="#EAF3FF",
            selectforeground="#202124",
            font=("Microsoft YaHei UI", 10),
        )
        self.listbox.pack(fill="both", expand=True)

        self._trace_id = self.variable.trace_add("write", self._on_text_changed)
        self.entry.bind("<FocusIn>", self._on_focus_in)
        self.entry.bind("<Down>", self._on_down)
        self.entry.bind("<Up>", self._on_up)
        self.entry.bind("<Return>", self._on_return)
        self.entry.bind("<Escape>", self._on_escape)
        self.entry.bind("<FocusOut>", self._on_focus_out)
        self.bind("<Configure>", lambda _e: self._position_popup(), add="+")
        self.popup_parent.bind("<Configure>", lambda _e: self.hide(), add="+")
        self.listbox.bind("<ButtonPress-1>", self._on_listbox_press)
        self.listbox.bind("<ButtonRelease-1>", self._on_listbox_release)
        self.listbox.bind("<Motion>", self._on_listbox_motion)

    def get(self):
        return self.variable.get()

    def set(self, value):
        self._set_value(value)

    def focus_set(self):
        self.entry.focus_set()

    def icursor(self, index):
        self.entry.icursor(index)

    def toggle_all(self):
        if self.popup.winfo_ismapped():
            self.hide()
            self.entry.focus_set()
            return
        suggestions = self.get_all_suggestions() or self.get_suggestions(self.variable.get()) or []
        self.show(suggestions)
        self.entry.focus_set()

    def show(self, suggestions):
        self.options = [self._normalize_item(item) for item in list(suggestions)[: self.max_items]]
        self.items = [item["value"] for item in self.options]
        self.listbox.delete(0, tk.END)
        if not self.options:
            self.hide()
            return
        for item in self.options:
            self.listbox.insert(tk.END, item["label"])
        self.listbox.configure(height=min(len(self.options), self.max_items))
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(0)
        self.listbox.activate(0)
        self._position_popup()
        self.popup.lift()
        self.entry.focus_set()

    def hide(self):
        if self._hide_after_id:
            self.after_cancel(self._hide_after_id)
            self._hide_after_id = None
        self.popup.place_forget()

    def _normalize_item(self, item):
        if isinstance(item, dict):
            label = item.get("label") or item.get("text") or item.get("display")
            value = item.get("value") or item.get("name") or item.get("contact") or label
            return {"label": str(label or value or ""), "value": str(value or ""), "raw": item.get("raw", item)}
        if isinstance(item, (tuple, list)) and item:
            label = item[0]
            value = item[1] if len(item) > 1 else item[0]
            raw = item[2] if len(item) > 2 else item
            return {"label": str(label or value or ""), "value": str(value or ""), "raw": raw}
        return {"label": str(item or ""), "value": str(item or ""), "raw": item}

    def _on_text_changed(self, *_args):
        if self._trace_busy:
            return
        text = self.variable.get()
        if self.entry.focus_get() != self.entry:
            return
        if len(text.strip()) < self.min_chars:
            self.hide()
            return
        suggestions = self.get_suggestions(text) or []
        self.show(suggestions)

    def _on_focus_in(self, _event=None):
        text = self.variable.get()
        if len(text.strip()) >= self.min_chars:
            suggestions = self.get_suggestions(text) or []
            if suggestions:
                self.show(suggestions)

    def _position_popup(self):
        if not self.options:
            return
        self.update_idletasks()
        root = self.popup_parent
        x = self.entry.winfo_rootx() - root.winfo_rootx()
        y = self.entry.winfo_rooty() - root.winfo_rooty() + self.entry.winfo_height()
        width = max(self.entry.winfo_width(), 1)
        row_height = 28
        height = min(len(self.options), self.max_items) * row_height + 2
        self.popup.place(x=x, y=y, width=width, height=height)

    def _move_selection(self, delta):
        if not self.items or not self.popup.winfo_ismapped():
            return False
        current = self.listbox.curselection()
        index = current[0] if current else 0
        index = max(0, min(len(self.items) - 1, index + delta))
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(index)
        self.listbox.activate(index)
        self.listbox.see(index)
        self.entry.focus_set()
        return True

    def _on_down(self, _event=None):
        if self._move_selection(1):
            return "break"
        suggestions = self.get_suggestions(self.variable.get()) or []
        self.show(suggestions)
        return "break" if self.items else None

    def _on_up(self, _event=None):
        if self._move_selection(-1):
            return "break"
        return None

    def _on_return(self, _event=None):
        if self.popup.winfo_ismapped() and self.items:
            self._choose_active()
            return "break"
        return None

    def _on_escape(self, _event=None):
        self.hide()
        return "break"

    def _on_listbox_press(self, _event=None):
        if self._hide_after_id:
            self.after_cancel(self._hide_after_id)
            self._hide_after_id = None

    def _on_listbox_release(self, _event=None):
        index = self.listbox.nearest(_event.y) if _event else None
        self._choose_index(index)
        return "break"

    def _on_listbox_motion(self, event):
        if not self.options:
            return
        index = self.listbox.nearest(event.y)
        if 0 <= index < len(self.options):
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(index)
            self.listbox.activate(index)

    def _choose_active(self):
        current = self.listbox.curselection()
        if not current:
            return
        self._choose_index(current[0])

    def _choose_index(self, index):
        if index is None or not (0 <= index < len(self.options)):
            return
        option = self.options[index]
        value = option["value"]
        self._set_value(value)
        self.hide()
        self.entry.icursor(tk.END)
        self.entry.focus_set()
        if self.on_select_item:
            self.on_select_item(option["raw"])
        if self.on_select:
            self.on_select(value)

    def _set_value(self, value):
        self._trace_busy = True
        try:
            self.variable.set(value)
        finally:
            self._trace_busy = False

    def _on_focus_out(self, _event=None):
        if self._hide_after_id:
            self.after_cancel(self._hide_after_id)
        self._hide_after_id = self.after(120, self._hide_if_focus_left)

    def _hide_if_focus_left(self):
        self._hide_after_id = None
        focus = self.focus_get()
        allowed = [self.entry, self.listbox]
        if self.button is not None:
            allowed.append(self.button)
        if focus not in allowed:
            self.hide()
