from __future__ import annotations

import logging
import shutil
from datetime import datetime
from pathlib import Path
from typing import Callable

from bizactivity import write_job_to_bizactivity
from excel_writer import write_header_block
from models import AppConfig, BizactivityResult, ClientRecord, OrderRequest, OrderResult, WriteResult
from sequence import (
    build_folder_job_number,
    build_order_folder_name,
    build_workbook_filename,
    detect_next_sequence,
    sanitize_job_description,
)
from storage import (
    get_client_by_id,
    get_client_by_name,
    next_internal_order_id,
    record_order_event,
)


class OrderCreationError(Exception):
    pass


def _resolve_client(req: OrderRequest, db_path: Path | None) -> ClientRecord:
    client: ClientRecord | None = None
    if req.client_id is not None:
        client = get_client_by_id(req.client_id, db_path=db_path)
    if client is None and req.client_name.strip():
        client = get_client_by_name(req.client_name.strip(), db_path=db_path)
    if client is None:
        raise OrderCreationError("Unknown client. Please reselect the client.")
    if client.status == "archived":
        raise OrderCreationError("Selected client is archived and cannot be used for new orders.")
    return client


def _resolve_template_path(config: AppConfig, template_id: str = "") -> Path:
    target_id = template_id.strip() or config.default_template_id
    selected = None
    for template in config.templates:
        if template.id == target_id:
            selected = template
            break
    if selected is None:
        for template in config.templates:
            if template.id == config.default_template_id:
                selected = template
                break
    if selected is None:
        selected = config.templates[0]

    source = Path(selected.source_path)
    if not source.is_absolute():
        source = Path(config.root_paths.templates_root) / source

    if not source.exists():
        raise OrderCreationError(f"Template file not found: {source}")
    return source


def create_order(
    req: OrderRequest,
    config: AppConfig,
    *,
    logger: logging.Logger | None = None,
    now_provider: Callable[[], datetime] = datetime.now,
    excel_write_func: Callable[..., WriteResult] = write_header_block,
    db_path: Path | None = None,
) -> OrderResult:
    client = _resolve_client(req, db_path)

    resolved_folder_path = client.folder_path.strip() or client.name
    client_folder = Path(resolved_folder_path)
    if not client_folder.is_absolute():
        client_folder = Path(config.root_paths.clients_root) / client_folder

    client_folder.mkdir(parents=True, exist_ok=True)

    sequence = detect_next_sequence(str(client_folder), client.abbr)
    description = sanitize_job_description(req.job_description or "")

    prefix = config.naming.order_prefix
    folder_name = build_order_folder_name(prefix, client.abbr, sequence, description)
    order_folder = client_folder / folder_name

    while order_folder.exists():
        sequence += 1
        folder_name = build_order_folder_name(prefix, client.abbr, sequence, description)
        order_folder = client_folder / folder_name

    order_folder.mkdir(parents=False, exist_ok=False)

    workbook_name = build_workbook_filename(prefix, client.abbr, sequence, ".xls")
    workbook_path = order_folder / workbook_name

    template_path = _resolve_template_path(config, client.template_id)
    shutil.copy2(template_path, workbook_path)

    folder_job_number = build_folder_job_number(client.abbr, sequence)
    internal_order_id = next_internal_order_id(now_provider(), db_path=db_path)

    excel_write_success = True
    excel_error_message = None
    excel_written_cells: list[str] = []

    if config.behavior_flags.allow_excel_write:
        mapping = {
            "sheet_name": config.excel_mapping.sheet_name,
            "fields": config.excel_mapping.fields,
        }
        values = {
            "client_name": client.name,
            "job_number": folder_job_number,
            "job_description": description,
            "due_date": req.due_date,
            "contact_person": client.contact_person,
            "phone": client.phone,
            "email": client.email,
            "street_address": client.street_address,
            "city_state_zip": client.city_state_zip,
        }
        write_result = excel_write_func(
            str(workbook_path),
            mapping,
            values,
            config.behavior_flags.skip_non_empty_cells,
            logger=logger,
        )
        excel_write_success = write_result.success
        excel_error_message = write_result.error_message
        excel_written_cells = write_result.written_cells

    # Best-effort write to bizactivity workbook
    bizactivity_success = True
    bizactivity_error: str | None = None
    if config.bizactivity_path:
        biz_values = {
            "client": client.name,
            "job_number": folder_job_number,
            "job_description": description,
            "create_date": now_provider().strftime("%Y-%m-%d"),
        }
        biz_result = write_job_to_bizactivity(
            config.bizactivity_path, biz_values,
            source_path=str(workbook_path),
            logger=logger,
        )
        bizactivity_success = biz_result.success
        bizactivity_error = biz_result.error_message
        if not biz_result.success and logger:
            logger.warning(
                "Bizactivity write failed for %s: %s",
                folder_job_number,
                biz_result.error_message,
            )

    record_order_event(
        internal_order_id=internal_order_id,
        client_id=client.id,
        client_name=client.name,
        folder_job_number=folder_job_number,
        folder_path=str(order_folder),
        workbook_path=str(workbook_path),
        db_path=db_path,
    )

    if logger:
        logger.info(
            "Created order folder=%s workbook=%s internal_order_id=%s client_id=%s excel_success=%s",
            order_folder,
            workbook_path,
            internal_order_id,
            client.id,
            excel_write_success,
        )

    return OrderResult(
        order_folder_path=str(order_folder),
        workbook_path=str(workbook_path),
        folder_job_number=folder_job_number,
        folder_name=folder_name,
        internal_order_id=internal_order_id,
        client_id=client.id,
        excel_write_success=excel_write_success,
        excel_error_message=excel_error_message,
        excel_written_cells=excel_written_cells,
        bizactivity_success=bizactivity_success,
        bizactivity_error=bizactivity_error,
    )
