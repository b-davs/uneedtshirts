from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class RootPaths:
    clients_root: str
    templates_root: str
    dashboard_root: str = ""
    orders_root_mode: str = "clients_root"


@dataclass
class NamingRules:
    order_prefix: str = "U"
    separator: str = "-"
    allow_description: bool = True
    order_folder_pattern: str = "U-{ABBR}-{SEQ}{DESC_PART}"
    workbook_filename_pattern: str = "U-{ABBR}-{SEQ}.xls"
    map_job_number_pattern: str = "{ABBR}-{SEQ}"


@dataclass
class BehaviorFlags:
    allow_excel_write: bool = True
    skip_non_empty_cells: bool = True
    auto_open_folder: bool = False
    auto_open_workbook: bool = False


@dataclass
class TemplateRecord:
    id: str
    label: str
    source_path: str
    dest_extension: str = ".xls"


@dataclass
class LegacyClientSeed:
    name: str
    abbr: str
    folder_name: str
    contact_person: str = ""
    phone: str = ""
    email: str = ""
    street_address: str = ""
    city_state_zip: str = ""


@dataclass
class ClientRecord:
    id: Optional[int] = None
    name: str = ""
    abbr: str = ""
    folder_path: str = ""
    contact_person: str = ""
    phone: str = ""
    email: str = ""
    street_address: str = ""
    city_state_zip: str = ""
    template_id: str = ""
    status: str = "active"
    created_at: str = ""
    updated_at: str = ""


@dataclass
class ExcelMapping:
    sheet_name: str = "Map"
    fields: dict[str, str] = field(default_factory=dict)


@dataclass
class AppConfig:
    root_paths: RootPaths
    naming: NamingRules
    behavior_flags: BehaviorFlags
    default_template_id: str
    templates: list[TemplateRecord]
    excel_mapping: ExcelMapping
    legacy_clients: list[LegacyClientSeed] = field(default_factory=list)
    bizactivity_path: str = ""


@dataclass
class OrderRequest:
    client_name: str = ""
    client_id: Optional[int] = None
    job_description: str = ""
    due_date: Optional[str] = None


@dataclass
class WriteResult:
    success: bool
    written_cells: list[str] = field(default_factory=list)
    skipped_cells: list[str] = field(default_factory=list)
    error_message: Optional[str] = None


@dataclass
class OrderResult:
    order_folder_path: str
    workbook_path: str
    folder_job_number: str
    folder_name: str
    internal_order_id: str
    client_id: Optional[int] = None
    excel_write_success: bool = True
    excel_error_message: Optional[str] = None
    excel_written_cells: list[str] = field(default_factory=list)
    bizactivity_success: bool = True
    bizactivity_error: Optional[str] = None


@dataclass
class BizactivityResult:
    success: bool
    action: str = ""  # "inserted", "updated", "moved", "skipped"
    target_row: Optional[int] = None
    month: Optional[int] = None
    error_message: Optional[str] = None


@dataclass
class SyncReport:
    synced: int = 0
    skipped: int = 0
    errors: int = 0
    details: list[str] = field(default_factory=list)

    def summary(self) -> str:
        return f"synced={self.synced}, skipped={self.skipped}, errors={self.errors}"


@dataclass
class SeedReport:
    created: int = 0
    updated: int = 0
    skipped: int = 0
    warnings: list[str] = field(default_factory=list)

    def summary(self) -> str:
        return (
            f"created={self.created}, updated={self.updated}, "
            f"skipped={self.skipped}, warnings={len(self.warnings)}"
        )
