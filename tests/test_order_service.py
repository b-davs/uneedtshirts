from __future__ import annotations

from datetime import datetime
from pathlib import Path

from models import (
    AppConfig,
    BehaviorFlags,
    ExcelMapping,
    NamingRules,
    OrderRequest,
    RootPaths,
    TemplateRecord,
    WriteResult,
)
from order_service import create_order
from storage import upsert_client
from models import ClientRecord


def _build_config(tmp_path: Path, template_path: Path) -> AppConfig:
    clients_root = tmp_path / "clients"
    clients_root.mkdir(parents=True, exist_ok=True)

    return AppConfig(
        root_paths=RootPaths(
            clients_root=str(clients_root),
            templates_root=str(template_path.parent),
            dashboard_root="",
            orders_root_mode="clients_root",
        ),
        naming=NamingRules(),
        behavior_flags=BehaviorFlags(
            allow_excel_write=True,
            skip_non_empty_cells=True,
            auto_open_folder=False,
            auto_open_workbook=False,
        ),
        default_template_id="whole_job_docs",
        templates=[
            TemplateRecord(
                id="whole_job_docs",
                label="Whole Job Docs",
                source_path=str(template_path),
                dest_extension=".xls",
            )
        ],
        excel_mapping=ExcelMapping(
            sheet_name="Map",
            fields={
                "client_name": "A2",
                "job_number": "B2",
                "job_description": "C2",
                "due_date": "E2",
                "contact_person": "AB2",
                "phone": "AC2",
                "email": "AD2",
                "street_address": "AE2",
                "city_state_zip": "AF2",
            },
        ),
        legacy_clients=[],
    )


def test_create_order_keeps_files_when_excel_write_fails(tmp_path: Path) -> None:
    template_path = tmp_path / "Whole Job Docs.xls"
    template_path.write_text("template", encoding="utf-8")

    config = _build_config(tmp_path, template_path)
    client = upsert_client(
        ClientRecord(
            name="Tamarac Fire Rescue",
            abbr="TFR",
            folder_path=str(Path(config.root_paths.clients_root) / "Tamarac Fire Rescue"),
            contact_person="Dispatcher",
        ),
        db_path=tmp_path / "state.db",
    )

    def fake_excel_fail(*args, **kwargs):
        return WriteResult(success=False, error_message="Workbook locked")

    result = create_order(
        OrderRequest(
            client_id=client.id,
            client_name=client.name,
            job_description="event shirts",
            due_date="2026-02-09",
        ),
        config,
        now_provider=lambda: datetime(2026, 2, 8, 10, 0, 0),
        excel_write_func=fake_excel_fail,
        db_path=tmp_path / "state.db",
    )

    assert Path(result.order_folder_path).exists()
    assert Path(result.workbook_path).exists()
    assert result.excel_write_success is False
    assert result.excel_error_message == "Workbook locked"
    assert result.folder_name == "U-TFR-1 event shirts"
    assert result.internal_order_id == "2602-0001"
    assert result.client_id == client.id


def test_create_order_passes_expected_excel_values(tmp_path: Path) -> None:
    template_path = tmp_path / "Whole Job Docs.xls"
    template_path.write_text("template", encoding="utf-8")

    config = _build_config(tmp_path, template_path)
    client = upsert_client(
        ClientRecord(
            name="Tamarac Fire Rescue",
            abbr="TFR",
            folder_path=str(Path(config.root_paths.clients_root) / "Tamarac Fire Rescue"),
            contact_person="Dispatch",
            phone="9540000000",
            email="dispatch@example.com",
            street_address="10600 Riverside Dr",
            city_state_zip="Coral Springs, FL 33071",
        ),
        db_path=tmp_path / "state.db",
    )

    captured: dict = {}

    def fake_excel_success(workbook_path, mapping, values, skip_non_empty, logger=None):
        captured["workbook_path"] = workbook_path
        captured["mapping"] = mapping
        captured["values"] = values
        captured["skip_non_empty"] = skip_non_empty
        return WriteResult(success=True, written_cells=["A2", "B2", "C2", "E2", "AB2"])

    result = create_order(
        OrderRequest(
            client_id=client.id,
            client_name=client.name,
            job_description="event shirts",
            due_date="2026-02-09",
        ),
        config,
        now_provider=lambda: datetime(2026, 2, 8, 10, 0, 0),
        excel_write_func=fake_excel_success,
        db_path=tmp_path / "state.db",
    )

    assert result.excel_write_success is True
    assert captured["values"]["client_name"] == "Tamarac Fire Rescue"
    assert captured["values"]["job_number"] == "TFR-1"
    assert captured["values"]["job_description"] == "event shirts"
    assert captured["values"]["due_date"] == "2026-02-09"
    assert captured["values"]["contact_person"] == "Dispatch"
    assert captured["values"]["phone"] == "9540000000"
    assert captured["values"]["email"] == "dispatch@example.com"
    assert captured["values"]["street_address"] == "10600 Riverside Dr"
    assert captured["values"]["city_state_zip"] == "Coral Springs, FL 33071"
    assert captured["mapping"]["sheet_name"] == "Map"
