from __future__ import annotations

import json
import sys
from pathlib import Path

from models import (
    AppConfig,
    BehaviorFlags,
    ExcelMapping,
    LegacyClientSeed,
    NamingRules,
    RootPaths,
    TemplateRecord,
)


class ConfigError(Exception):
    pass


def runtime_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _load_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise ConfigError(f"Config file not found: {path}") from exc
    except json.JSONDecodeError as exc:
        raise ConfigError(f"Invalid JSON in {path}: {exc}") from exc


def _build_legacy_seed_clients(raw_clients: list[dict]) -> list[LegacyClientSeed]:
    legacy: list[LegacyClientSeed] = []
    for item in raw_clients:
        name = str(item.get("name", "")).strip()
        if not name:
            continue

        city_state_zip = ""
        city = str(item.get("address_city", "")).strip()
        state = str(item.get("address_state", "")).strip()
        zip_code = str(item.get("address_zip", "")).strip()
        if city and state:
            city_state_zip = f"{city}, {state}"
            if zip_code:
                city_state_zip = f"{city_state_zip} {zip_code}"

        legacy.append(
            LegacyClientSeed(
                name=name,
                abbr=str(item.get("abbr", "")).strip().upper(),
                folder_name=str(item.get("folder_name", "")).strip() or name,
                contact_person=str(item.get("contact_person", "")).strip(),
                phone=str(item.get("phone", "")).strip(),
                email=str(item.get("email", "")).strip(),
                street_address=str(item.get("address_street", "")).strip(),
                city_state_zip=city_state_zip,
            )
        )
    return legacy


def _to_config(raw: dict) -> AppConfig:
    try:
        root = RootPaths(**raw["root_paths"])
        naming = NamingRules(**raw["naming"])
        behavior = BehaviorFlags(**raw["behavior_flags"])
        templates = [TemplateRecord(**item) for item in raw["templates"]]
        excel_mapping = ExcelMapping(**raw["excel_mapping"])
        default_template_id = raw["default_template_id"]
    except KeyError as exc:
        raise ConfigError(f"Missing config key: {exc}") from exc
    except TypeError as exc:
        raise ConfigError(f"Invalid config object shape: {exc}") from exc

    if not templates:
        raise ConfigError("At least one template is required in config.")

    legacy_clients = _build_legacy_seed_clients(raw.get("clients", []))

    return AppConfig(
        root_paths=root,
        naming=naming,
        behavior_flags=behavior,
        default_template_id=default_template_id,
        templates=templates,
        excel_mapping=excel_mapping,
        legacy_clients=legacy_clients,
    )


def load_runtime_config(base_dir: Path | None = None) -> tuple[AppConfig, Path]:
    base = base_dir or runtime_base_dir()
    config_path = base / "config.json"
    example_path = base / "config.example.json"

    source = config_path if config_path.exists() else example_path
    if not source.exists():
        raise ConfigError(
            "No config found. Expected config.json or config.example.json in app directory."
        )

    config = _to_config(_load_json(source))
    return config, config_path
