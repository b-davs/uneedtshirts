from __future__ import annotations

import logging
import os
import sqlite3
import subprocess
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, simpledialog, ttk

from excel_writer import write_header_block
from models import AppConfig, ClientRecord, OrderRequest, OrderResult
from order_service import OrderCreationError, create_order
from sequence import sanitize_job_description
from storage import (
    add_job_description,
    delete_job_description,
    get_client_by_id,
    list_clients,
    list_job_description_records,
    list_job_descriptions,
    update_job_description,
    upsert_client,
)
from ui_manage_clients import open_manage_clients
from ui_new_client import prompt_client_editor


class ManageJobDescriptionsDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, client: ClientRecord) -> None:
        super().__init__(parent)
        self.title(f"Manage Job Descriptions — {client.name}")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self._client = client
        self.changed = False

        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)

        self._listbox = tk.Listbox(frame, width=50, height=12)
        self._listbox.grid(row=0, column=0, rowspan=4, padx=(0, 8))

        ttk.Button(frame, text="Add", command=self._add).grid(row=0, column=1, sticky="ew", pady=2)
        ttk.Button(frame, text="Edit", command=self._edit).grid(row=1, column=1, sticky="ew", pady=2)
        ttk.Button(frame, text="Delete", command=self._delete).grid(row=2, column=1, sticky="ew", pady=2)
        ttk.Button(frame, text="Close", command=self.destroy).grid(row=3, column=1, sticky="ew", pady=2)

        self._records: list[tuple[int, str]] = []
        self._reload()

        self.bind("<Escape>", lambda _: self.destroy())

    def _reload(self) -> None:
        self._records = list_job_description_records(self._client.id)
        self._listbox.delete(0, tk.END)
        for _, name in self._records:
            self._listbox.insert(tk.END, name)

    def _add(self) -> None:
        name = simpledialog.askstring("Add Job Description", "Name:", parent=self)
        if not name or not name.strip():
            return
        try:
            add_job_description(self._client.id, name.strip())
            self.changed = True
            self._reload()
        except sqlite3.IntegrityError:
            messagebox.showerror("Duplicate", "That job description already exists.", parent=self)

    def _edit(self) -> None:
        sel = self._listbox.curselection()
        if not sel:
            messagebox.showinfo("Select", "Select an item to edit.", parent=self)
            return
        idx = sel[0]
        desc_id, old_name = self._records[idx]
        new_name = simpledialog.askstring(
            "Edit Job Description", "Name:", initialvalue=old_name, parent=self
        )
        if not new_name or not new_name.strip() or new_name.strip() == old_name:
            return
        try:
            update_job_description(desc_id, new_name.strip())
            self.changed = True
            self._reload()
        except sqlite3.IntegrityError:
            messagebox.showerror("Duplicate", "That job description already exists.", parent=self)

    def _delete(self) -> None:
        sel = self._listbox.curselection()
        if not sel:
            messagebox.showinfo("Select", "Select an item to delete.", parent=self)
            return
        idx = sel[0]
        desc_id, name = self._records[idx]
        if not messagebox.askyesno("Confirm Delete", f"Delete \"{name}\"?", parent=self):
            return
        delete_job_description(desc_id)
        self.changed = True
        self._reload()


class MainWindow:
    NEW_CLIENT_LABEL = "+ New Client..."

    def __init__(
        self,
        root: tk.Tk,
        config: AppConfig,
        config_path: Path,
        logger: logging.Logger,
    ) -> None:
        self.root = root
        self.config = config
        self.config_path = config_path
        self.base_dir = config_path.parent
        self.logger = logger
        self.last_result: OrderResult | None = None
        self.last_request: OrderRequest | None = None
        self.client_by_name: dict[str, ClientRecord] = {}
        self._desc_is_dropdown = False

        root.title("New Order Launcher")
        root.geometry("780x300")

        self.client_var = tk.StringVar()
        self.description_var = tk.StringVar()
        self.due_date_var = tk.StringVar()

        self._build_ui()
        self._reload_clients()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill="both", expand=True)
        self._main_frame = frame

        ttk.Label(frame, text="Client *").grid(row=0, column=0, sticky="w", pady=6)
        self.client_combo = ttk.Combobox(
            frame,
            textvariable=self.client_var,
            state="readonly",
            width=42,
        )
        self.client_combo.grid(row=0, column=1, sticky="ew", pady=6)
        self.client_combo.bind("<<ComboboxSelected>>", self._on_client_selected)

        ttk.Button(frame, text="Manage Clients...", command=self._open_manage_clients).grid(
            row=0, column=2, padx=(8, 0), sticky="w"
        )

        # Row 1: Job Description — label + freeform entry (default)
        self._desc_label = ttk.Label(frame, text="Job Description")
        self._desc_label.grid(row=1, column=0, sticky="w", pady=6)

        self._desc_entry = ttk.Entry(frame, textvariable=self.description_var, width=45)
        self._desc_entry.grid(row=1, column=1, sticky="ew", pady=6)

        # Row 1 alternates: dropdown combo + manage button (hidden initially)
        self._desc_combo = ttk.Combobox(
            frame,
            textvariable=self.description_var,
            state="readonly",
            width=42,
        )
        self._desc_manage_btn = ttk.Button(
            frame, text="Manage...", command=self._open_manage_job_descriptions
        )

        ttk.Label(frame, text="Due Date").grid(row=2, column=0, sticky="w", pady=6)
        self.due_date_widget = self._build_due_date_widget(frame)
        self.due_date_widget.grid(row=2, column=1, sticky="w", pady=6)

        ttk.Button(frame, text="Create Order", command=self._create_order).grid(
            row=3, column=1, sticky="w", pady=(14, 6)
        )

        actions = ttk.Frame(frame)
        actions.grid(row=4, column=1, sticky="w", pady=(6, 0))

        self.open_folder_btn = ttk.Button(
            actions,
            text="Open Folder",
            command=self._open_folder,
            state="disabled",
        )
        self.open_folder_btn.grid(row=0, column=0, padx=(0, 6))

        self.open_workbook_btn = ttk.Button(
            actions,
            text="Open Workbook",
            command=self._open_workbook,
            state="disabled",
        )
        self.open_workbook_btn.grid(row=0, column=1, padx=(0, 6))

        self.retry_excel_btn = ttk.Button(
            actions,
            text="Retry Excel Write",
            command=self._retry_excel_write,
            state="disabled",
        )
        self.retry_excel_btn.grid(row=0, column=2)

        frame.columnconfigure(1, weight=1)

    def _show_desc_entry(self, show_manage: bool = False) -> None:
        self._desc_combo.grid_remove()
        self._desc_label.configure(text="Job Description")
        self._desc_entry.grid(row=1, column=1, sticky="ew", pady=6)
        if show_manage:
            self._desc_manage_btn.grid(row=1, column=2, padx=(8, 0), sticky="w")
        else:
            self._desc_manage_btn.grid_remove()
        self.description_var.set("")
        self._desc_is_dropdown = False

    def _show_desc_dropdown(self, descriptions: list[str]) -> None:
        self._desc_entry.grid_remove()
        self._desc_label.configure(text="Job Description *")
        self._desc_combo["values"] = descriptions
        self._desc_combo.grid(row=1, column=1, sticky="ew", pady=6)
        self._desc_manage_btn.grid(row=1, column=2, padx=(8, 0), sticky="w")
        self.description_var.set("")
        self._desc_is_dropdown = True

    def _build_due_date_widget(self, parent: ttk.Frame) -> tk.Widget:
        try:
            from tkcalendar import DateEntry  # type: ignore[import-not-found]

            widget = DateEntry(
                parent,
                textvariable=self.due_date_var,
                date_pattern="yyyy-mm-dd",
                width=15,
            )
            return widget
        except Exception:
            return ttk.Entry(parent, textvariable=self.due_date_var, width=18)

    def _reload_clients(self) -> None:
        clients = list_clients(include_archived=False)
        self.client_by_name = {client.name: client for client in clients}
        values = sorted(self.client_by_name.keys()) + [self.NEW_CLIENT_LABEL]
        self.client_combo["values"] = values

    def _on_client_selected(self, _: object) -> None:
        selected = self.client_var.get()
        if selected == self.NEW_CLIENT_LABEL:
            new_client = self._prompt_add_client()
            if new_client is None:
                self.client_var.set("")
                self._show_desc_entry()
                return
            self.client_var.set(new_client.name)
            selected = new_client.name

        self._update_desc_field_for_client(selected)

    def _update_desc_field_for_client(self, client_name: str) -> None:
        client = self.client_by_name.get(client_name)
        if client and client.id is not None:
            descs = list_job_descriptions(client.id)
            if descs:
                self._show_desc_dropdown(descs)
                return
            # No descriptions yet — show freeform but keep Manage button visible
            # so the user can add their first job descriptions
            self._show_desc_entry(show_manage=True)
            return
        self._show_desc_entry()

    def _prompt_add_client(self) -> ClientRecord | None:
        default_path = str(Path(self.config.root_paths.clients_root) / "")
        new_client = prompt_client_editor(
            self.root,
            title="Add Client",
            initial_client=None,
            default_folder_path=default_path,
            templates=self.config.templates,
        )
        if new_client is None:
            return None

        try:
            persisted = upsert_client(new_client)
            Path(persisted.folder_path).mkdir(parents=True, exist_ok=True)
        except sqlite3.IntegrityError as exc:
            messagebox.showerror("Save Failed", f"Unique constraint failed: {exc}")
            return None
        except Exception as exc:
            self.logger.exception("Failed to save new client")
            messagebox.showerror("Save Failed", str(exc))
            return None

        self._reload_clients()
        return persisted

    def _open_manage_clients(self) -> None:
        changed = open_manage_clients(self.root, self.config, self.logger, self.base_dir)
        if changed:
            self._reload_clients()

    def _open_manage_job_descriptions(self) -> None:
        selected_name = self.client_var.get().strip()
        client = self.client_by_name.get(selected_name)
        if client is None or client.id is None:
            return

        dialog = ManageJobDescriptionsDialog(self.root, client)
        self.root.wait_window(dialog)

        if dialog.changed:
            descs = list_job_descriptions(client.id)
            if descs:
                self._show_desc_dropdown(descs)
            else:
                self._show_desc_entry()

    def _create_order(self) -> None:
        selected_name = self.client_var.get().strip()
        if not selected_name or selected_name == self.NEW_CLIENT_LABEL:
            messagebox.showerror("Validation", "Client is required.")
            return

        selected_client = self.client_by_name.get(selected_name)
        if selected_client is None:
            messagebox.showerror("Validation", "Client not found. Please reselect.")
            return

        job_description = self.description_var.get().strip()
        if self._desc_is_dropdown and not job_description:
            messagebox.showerror("Validation", "Job Description is required.")
            return

        due_date = self.due_date_var.get().strip() or None
        req = OrderRequest(
            client_id=selected_client.id,
            client_name=selected_client.name,
            job_description=job_description,
            due_date=due_date,
        )
        self.last_request = req

        try:
            result = create_order(req, self.config, logger=self.logger)
        except OrderCreationError as exc:
            messagebox.showerror("Order Creation Failed", str(exc))
            return
        except Exception as exc:
            self.logger.exception("Unexpected order creation failure")
            messagebox.showerror("Order Creation Failed", str(exc))
            return

        self.last_result = result
        self._set_result_actions_enabled(True)
        self.retry_excel_btn.configure(
            state="normal" if not result.excel_write_success else "disabled"
        )

        if result.excel_write_success:
            messagebox.showinfo(
                "Order Created",
                f"Order created successfully.\n\nFolder: {result.order_folder_path}\nWorkbook: {result.workbook_path}",
            )
        else:
            messagebox.showwarning(
                "Order Created (Excel Write Failed)",
                "Folder and workbook were created, but writing to Excel failed.\n"
                f"Error: {result.excel_error_message}",
            )

    def _retry_excel_write(self) -> None:
        if not self.last_result or not self.last_request:
            return

        client = None
        if self.last_result.client_id is not None:
            client = get_client_by_id(self.last_result.client_id)
        if client is None and self.last_request.client_name:
            client = self.client_by_name.get(self.last_request.client_name)
        if client is None:
            messagebox.showerror("Retry Failed", "Client not found in client registry.")
            return

        values = {
            "client_name": client.name,
            "job_number": self.last_result.folder_job_number,
            "job_description": sanitize_job_description(self.last_request.job_description),
            "due_date": self.last_request.due_date,
            "contact_person": client.contact_person,
            "phone": client.phone,
            "email": client.email,
            "street_address": client.street_address,
            "city_state_zip": client.city_state_zip,
        }
        mapping = {
            "sheet_name": self.config.excel_mapping.sheet_name,
            "fields": self.config.excel_mapping.fields,
        }
        result = write_header_block(
            self.last_result.workbook_path,
            mapping,
            values,
            self.config.behavior_flags.skip_non_empty_cells,
            logger=self.logger,
        )

        if result.success:
            self.retry_excel_btn.configure(state="disabled")
            messagebox.showinfo("Retry Succeeded", "Excel write completed successfully.")
        else:
            messagebox.showerror(
                "Retry Failed",
                f"Excel write failed again.\nError: {result.error_message}",
            )

    def _set_result_actions_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.open_folder_btn.configure(state=state)
        self.open_workbook_btn.configure(state=state)

    def _open_folder(self) -> None:
        if self.last_result:
            self._open_path(self.last_result.order_folder_path)

    def _open_workbook(self) -> None:
        if self.last_result:
            self._open_path(self.last_result.workbook_path)

    def _open_path(self, path: str) -> None:
        try:
            if hasattr(os, "startfile"):
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                subprocess.run(["open", path], check=False)
        except Exception:
            self.logger.exception("Failed to open path: %s", path)
            messagebox.showerror("Open Failed", f"Could not open: {path}")
