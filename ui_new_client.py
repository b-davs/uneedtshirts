from __future__ import annotations

import re
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from models import ClientRecord, TemplateRecord

ABBR_PATTERN = re.compile(r"^[A-Z0-9]+$")

DEFAULT_TEMPLATE_LABEL = "(Default)"


class ClientEditorDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        *,
        title: str,
        initial_client: ClientRecord | None,
        default_folder_path: str,
        templates: list[TemplateRecord] | None = None,
    ) -> None:
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self._initial_client = initial_client
        self._default_folder_path = default_folder_path
        self._templates = templates or []
        self.result: ClientRecord | None = None

        self.vars: dict[str, tk.StringVar] = {
            "name": tk.StringVar(value=(initial_client.name if initial_client else "")),
            "abbr": tk.StringVar(value=(initial_client.abbr if initial_client else "")),
            "folder_path": tk.StringVar(
                value=(
                    initial_client.folder_path
                    if initial_client and initial_client.folder_path
                    else default_folder_path
                )
            ),
            "contact_person": tk.StringVar(
                value=(initial_client.contact_person if initial_client else "")
            ),
            "phone": tk.StringVar(value=(initial_client.phone if initial_client else "")),
            "email": tk.StringVar(value=(initial_client.email if initial_client else "")),
            "street_address": tk.StringVar(
                value=(initial_client.street_address if initial_client else "")
            ),
            "city_state_zip": tk.StringVar(
                value=(initial_client.city_state_zip if initial_client else "")
            ),
        }

        self._template_var = tk.StringVar()
        self._template_label_to_id: dict[str, str] = {}
        self._setup_template_choices(initial_client)

        labels = [
            ("Client Name *", "name"),
            ("Abbreviation *", "abbr"),
            ("Folder Path *", "folder_path"),
            ("Contact Person", "contact_person"),
            ("Phone", "phone"),
            ("Email", "email"),
            ("Street Address", "street_address"),
            ("City/State/Zip", "city_state_zip"),
        ]

        frame = ttk.Frame(self, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")

        for idx, (label_text, key) in enumerate(labels):
            ttk.Label(frame, text=label_text).grid(row=idx, column=0, sticky="w", pady=2)
            entry = ttk.Entry(frame, textvariable=self.vars[key], width=54)
            entry.grid(row=idx, column=1, sticky="ew", pady=2)

        template_row = len(labels)
        ttk.Label(frame, text="Template").grid(row=template_row, column=0, sticky="w", pady=2)
        template_values = list(self._template_label_to_id.keys())
        self._template_combo = ttk.Combobox(
            frame,
            textvariable=self._template_var,
            state="readonly",
            values=template_values,
            width=51,
        )
        self._template_combo.grid(row=template_row, column=1, sticky="ew", pady=2)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=template_row + 1, column=0, columnspan=2, pady=(10, 0), sticky="e")

        ttk.Button(button_frame, text="Cancel", command=self._cancel).grid(
            row=0, column=0, padx=4
        )
        ttk.Button(button_frame, text="Save", command=self._save).grid(
            row=0, column=1, padx=4
        )

        self.bind("<Return>", lambda _: self._save())
        self.bind("<Escape>", lambda _: self._cancel())

    def _setup_template_choices(self, initial_client: ClientRecord | None) -> None:
        self._template_label_to_id[DEFAULT_TEMPLATE_LABEL] = ""
        for tmpl in self._templates:
            self._template_label_to_id[tmpl.label] = tmpl.id

        initial_template_id = initial_client.template_id if initial_client else ""
        selected_label = DEFAULT_TEMPLATE_LABEL
        if initial_template_id:
            for tmpl in self._templates:
                if tmpl.id == initial_template_id:
                    selected_label = tmpl.label
                    break
        self._template_var.set(selected_label)

    def _get_selected_template_id(self) -> str:
        label = self._template_var.get()
        return self._template_label_to_id.get(label, "")

    def _save(self) -> None:
        name = self.vars["name"].get().strip()
        abbr = self.vars["abbr"].get().strip().upper()
        folder_path = self.vars["folder_path"].get().strip()

        if not name:
            messagebox.showerror("Validation", "Client Name is required.", parent=self)
            return
        if not abbr:
            messagebox.showerror("Validation", "Abbreviation is required.", parent=self)
            return
        if not ABBR_PATTERN.match(abbr):
            messagebox.showerror(
                "Validation",
                "Abbreviation must contain only uppercase letters and numbers.",
                parent=self,
            )
            return
        if not folder_path:
            messagebox.showerror("Validation", "Folder Path is required.", parent=self)
            return

        initial = self._initial_client
        if initial is None:
            normalized_default = self._default_folder_path.rstrip("\\/")
            normalized_folder = folder_path.rstrip("\\/")
            if normalized_default and normalized_folder == normalized_default:
                folder_path = str(Path(folder_path) / name)
        self.result = ClientRecord(
            id=initial.id if initial else None,
            name=name,
            abbr=abbr,
            folder_path=folder_path,
            contact_person=self.vars["contact_person"].get().strip(),
            phone=self.vars["phone"].get().strip(),
            email=self.vars["email"].get().strip(),
            street_address=self.vars["street_address"].get().strip(),
            city_state_zip=self.vars["city_state_zip"].get().strip(),
            template_id=self._get_selected_template_id(),
            status=initial.status if initial else "active",
            created_at=initial.created_at if initial else "",
            updated_at=initial.updated_at if initial else "",
        )
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


def prompt_client_editor(
    parent: tk.Misc,
    *,
    title: str,
    initial_client: ClientRecord | None,
    default_folder_path: str,
    templates: list[TemplateRecord] | None = None,
) -> ClientRecord | None:
    dialog = ClientEditorDialog(
        parent,
        title=title,
        initial_client=initial_client,
        default_folder_path=default_folder_path,
        templates=templates,
    )
    parent.wait_window(dialog)
    return dialog.result
