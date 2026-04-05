from __future__ import annotations

import logging
import sqlite3
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from models import AppConfig, ClientRecord
from storage import (
    archive_client,
    get_client_by_id,
    has_orders_for_client,
    import_clients_from_csv,
    list_clients,
    upsert_client,
)
from ui_new_client import prompt_client_editor


class ManageClientsDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        config: AppConfig,
        logger: logging.Logger,
        base_dir: Path,
    ) -> None:
        super().__init__(parent)
        self.title("Manage Clients")
        self.geometry("980x420")
        self.transient(parent)
        self.grab_set()

        self.config = config
        self.logger = logger
        self.base_dir = base_dir
        self.changed = False
        self.show_archived_var = tk.BooleanVar(value=False)

        self._build_ui()
        self._refresh()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self, padding=10)
        frame.pack(fill="both", expand=True)

        controls = ttk.Frame(frame)
        controls.pack(fill="x")

        ttk.Button(controls, text="Add", command=self._add_client).pack(side="left", padx=3)
        ttk.Button(controls, text="Edit", command=self._edit_client).pack(side="left", padx=3)
        ttk.Button(controls, text="Archive", command=self._archive_client).pack(side="left", padx=3)
        ttk.Button(controls, text="Import CSV", command=self._import_csv).pack(
            side="left", padx=3
        )
        ttk.Checkbutton(
            controls,
            text="Show Archived",
            variable=self.show_archived_var,
            command=self._refresh,
        ).pack(side="right")

        columns = (
            "name",
            "abbr",
            "folder_path",
            "contact_person",
            "phone",
            "email",
            "status",
        )
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=14)
        self.tree.pack(fill="both", expand=True, pady=(10, 0))

        headings = {
            "name": "Name",
            "abbr": "Abbr",
            "folder_path": "Folder Path",
            "contact_person": "Contact",
            "phone": "Phone",
            "email": "Email",
            "status": "Status",
        }
        widths = {
            "name": 170,
            "abbr": 80,
            "folder_path": 250,
            "contact_person": 130,
            "phone": 120,
            "email": 180,
            "status": 90,
        }

        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w")

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(8, 0))
        ttk.Button(buttons, text="Close", command=self._close).pack(side="right")

    def _close(self) -> None:
        self.destroy()

    def _refresh(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        clients = list_clients(include_archived=self.show_archived_var.get())
        for client in clients:
            if client.id is None:
                continue
            self.tree.insert(
                "",
                "end",
                iid=str(client.id),
                values=(
                    client.name,
                    client.abbr,
                    client.folder_path,
                    client.contact_person,
                    client.phone,
                    client.email,
                    client.status,
                ),
            )

    def _selected_client(self) -> ClientRecord | None:
        selected = self.tree.selection()
        if not selected:
            return None
        client_id = int(selected[0])
        return get_client_by_id(client_id)

    def _add_client(self) -> None:
        default_path = str(Path(self.config.root_paths.clients_root) / "")
        created = prompt_client_editor(
            self,
            title="Add Client",
            initial_client=None,
            default_folder_path=default_path,
            templates=self.config.templates,
        )
        if created is None:
            return

        try:
            persisted = upsert_client(created)
            Path(persisted.folder_path).mkdir(parents=True, exist_ok=True)
        except sqlite3.IntegrityError as exc:
            messagebox.showerror("Save Failed", f"Unique constraint failed: {exc}", parent=self)
            return
        except Exception as exc:
            self.logger.exception("Failed to add client")
            messagebox.showerror("Save Failed", str(exc), parent=self)
            return

        self.changed = True
        self._refresh()

    def _edit_client(self) -> None:
        client = self._selected_client()
        if client is None:
            messagebox.showinfo("Edit Client", "Select a client first.", parent=self)
            return

        if client.id is not None and has_orders_for_client(client.id):
            proceed = messagebox.askyesno(
                "Edit Client",
                "This client already has orders. Historical folders will not change.\n"
                "Future orders will use new values. Continue?",
                parent=self,
            )
            if not proceed:
                return

        updated = prompt_client_editor(
            self,
            title="Edit Client",
            initial_client=client,
            default_folder_path=client.folder_path,
            templates=self.config.templates,
        )
        if updated is None:
            return

        try:
            persisted = upsert_client(updated)
            Path(persisted.folder_path).mkdir(parents=True, exist_ok=True)
        except sqlite3.IntegrityError as exc:
            messagebox.showerror("Save Failed", f"Unique constraint failed: {exc}", parent=self)
            return
        except Exception as exc:
            self.logger.exception("Failed to edit client")
            messagebox.showerror("Save Failed", str(exc), parent=self)
            return

        self.changed = True
        self._refresh()

    def _archive_client(self) -> None:
        client = self._selected_client()
        if client is None or client.id is None:
            messagebox.showinfo("Archive Client", "Select a client first.", parent=self)
            return

        confirm = messagebox.askyesno(
            "Archive Client",
            f"Archive client '{client.name}'?",
            parent=self,
        )
        if not confirm:
            return

        archive_client(client.id)
        self.changed = True
        self._refresh()

    def _import_csv(self) -> None:
        csv_path = self.base_dir / "clients.csv"
        report = import_clients_from_csv(self.config, csv_path)
        self.logger.info("CSV import report: %s", report.summary())
        if report.warnings:
            for warning in report.warnings[:20]:
                self.logger.warning(warning)

        self.changed = True
        self._refresh()
        messagebox.showinfo(
            "CSV Import",
            "Client import complete.\n\n"
            f"Created: {report.created}\n"
            f"Updated: {report.updated}\n"
            f"Skipped: {report.skipped}\n"
            f"Warnings: {len(report.warnings)}",
            parent=self,
        )


def open_manage_clients(
    parent: tk.Misc,
    config: AppConfig,
    logger: logging.Logger,
    base_dir: Path,
) -> bool:
    dialog = ManageClientsDialog(parent, config, logger, base_dir)
    parent.wait_window(dialog)
    return dialog.changed
