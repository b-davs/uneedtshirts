from pathlib import Path

from config import load_runtime_config


SAMPLE_CONFIG = """
{
  "root_paths": {
    "clients_root": "D:/A Client Sites & Images",
    "templates_root": "C:/Users/dan/Desktop",
    "dashboard_root": "",
    "orders_root_mode": "clients_root"
  },
  "naming": {
    "order_prefix": "U",
    "separator": "-",
    "allow_description": true,
    "order_folder_pattern": "U-{ABBR}-{SEQ}{DESC_PART}",
    "workbook_filename_pattern": "U-{ABBR}-{SEQ}.xls",
    "map_job_number_pattern": "{ABBR}-{SEQ}"
  },
  "behavior_flags": {
    "allow_excel_write": true,
    "skip_non_empty_cells": true,
    "auto_open_folder": false,
    "auto_open_workbook": false
  },
  "default_template_id": "whole_job_docs",
  "templates": [
    {
      "id": "whole_job_docs",
      "label": "Whole Job Docs",
      "source_path": "C:/Users/dan/Desktop/Whole Job Docs.xls",
      "dest_extension": ".xls"
    }
  ],
  "excel_mapping": {
    "sheet_name": "Map",
    "fields": {
      "client_name": "A2",
      "job_number": "B2",
      "job_description": "C2",
      "due_date": "E2",
      "contact_person": "AB2",
      "phone": "AC2",
      "email": "AD2",
      "street_address": "AE2",
      "city_state_zip": "AF2"
    }
  },
  "clients": [
    {
      "name": "Tamarac Fire Rescue",
      "abbr": "TFR",
      "folder_name": "Tamarac Fire Rescue",
      "contact_person": "Dispatcher",
      "phone": "9540000000",
      "email": "",
      "address_street": "123 Main St",
      "address_city": "Coral Springs",
      "address_state": "FL",
      "address_zip": "33071"
    }
  ]
}
"""


def test_load_runtime_config_from_config_json(tmp_path: Path) -> None:
    (tmp_path / "config.json").write_text(SAMPLE_CONFIG, encoding="utf-8")
    config, config_path = load_runtime_config(tmp_path)

    assert config.root_paths.clients_root == "D:/A Client Sites & Images"
    assert config.default_template_id == "whole_job_docs"
    assert config.excel_mapping.fields["city_state_zip"] == "AF2"
    assert config_path == tmp_path / "config.json"


def test_load_runtime_config_parses_legacy_clients(tmp_path: Path) -> None:
    (tmp_path / "config.json").write_text(SAMPLE_CONFIG, encoding="utf-8")
    config, _ = load_runtime_config(tmp_path)

    assert len(config.legacy_clients) == 1
    legacy = config.legacy_clients[0]
    assert legacy.name == "Tamarac Fire Rescue"
    assert legacy.abbr == "TFR"
    assert legacy.city_state_zip == "Coral Springs, FL 33071"
