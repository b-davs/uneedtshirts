from pathlib import Path

from models import AppConfig, BehaviorFlags, ExcelMapping, NamingRules, RootPaths, TemplateRecord
from storage import (
    archive_client,
    generate_client_abbreviation,
    has_orders_for_client,
    import_clients_from_csv,
    list_clients,
    parse_address_components,
    seed_clients_from_csv_if_needed,
    upsert_client,
)
from models import ClientRecord


def _build_config(tmp_path: Path) -> AppConfig:
    clients_root = tmp_path / "clients"
    clients_root.mkdir(parents=True, exist_ok=True)

    template_path = tmp_path / "Whole Job Docs.xls"
    template_path.write_text("template", encoding="utf-8")

    return AppConfig(
        root_paths=RootPaths(
            clients_root=str(clients_root),
            templates_root=str(tmp_path),
        ),
        naming=NamingRules(),
        behavior_flags=BehaviorFlags(),
        default_template_id="whole_job_docs",
        templates=[
            TemplateRecord(
                id="whole_job_docs",
                label="Whole Job Docs",
                source_path=str(template_path),
            )
        ],
        excel_mapping=ExcelMapping(fields={}),
        legacy_clients=[],
    )


def test_parse_address_components_with_suite() -> None:
    street, city_state_zip, warning = parse_address_components(
        "2015 SW 20th St #106, Ft. Lauderdale, FL 33315"
    )
    assert warning is False
    assert "#106" in street
    assert city_state_zip == "Ft. Lauderdale, FL 33315"


def test_parse_address_components_fallback_for_uncertain_values() -> None:
    street, city_state_zip, warning = parse_address_components("Unknown Place")
    assert warning is True
    assert street == "Unknown Place"
    assert city_state_zip == ""


def test_generate_client_abbreviation_collision() -> None:
    taken = {"TFR", "TFR2"}
    result = generate_client_abbreviation("Tamarac Fire Rescue", taken)
    assert result.startswith("TFR")
    assert result not in taken


def test_seed_clients_from_csv_if_needed_and_upsert(tmp_path: Path) -> None:
    config = _build_config(tmp_path)
    csv_path = tmp_path / "clients.csv"
    csv_path.write_text(
        "Client name,Contact person,phone number,client address,email\n"
        "Cintas,Mike,954-111-2222,\"2701 SW 145th Ave #270, Miramar, FL 33027\",mike@example.com\n"
        "Encore Dance,Anne,954-333-4444,\"3461 Hiatus Rd, Sunrise, FL 33351\",encore@example.com\n",
        encoding="utf-8",
    )

    report = seed_clients_from_csv_if_needed(config, csv_path, db_path=tmp_path / "state.db")

    assert report.created == 2
    assert report.updated == 0
    clients = list_clients(db_path=tmp_path / "state.db")
    assert len(clients) == 2


def test_import_clients_csv_updates_existing_by_name(tmp_path: Path) -> None:
    config = _build_config(tmp_path)
    db_path = tmp_path / "state.db"

    upsert_client(
        ClientRecord(
            name="Cintas",
            abbr="CINTAS",
            folder_path=str(Path(config.root_paths.clients_root) / "Cintas"),
            contact_person="Old",
        ),
        db_path=db_path,
    )

    csv_path = tmp_path / "clients.csv"
    csv_path.write_text(
        "Client name,Contact person,phone number,client address,email\n"
        "Cintas,New Contact,954-000-0000,\"2701 SW 145th Ave #270, Miramar, FL 33027\",new@example.com\n",
        encoding="utf-8",
    )

    report = import_clients_from_csv(config, csv_path, db_path=db_path)
    assert report.updated == 1

    clients = list_clients(db_path=db_path)
    assert clients[0].contact_person == "New Contact"
    assert clients[0].abbr == "CINTAS"


def test_archive_filter_and_order_link_check(tmp_path: Path) -> None:
    config = _build_config(tmp_path)
    db_path = tmp_path / "state.db"

    client = upsert_client(
        ClientRecord(
            name="Taravella GSA",
            abbr="TGSA",
            folder_path=str(Path(config.root_paths.clients_root) / "Taravella GSA"),
        ),
        db_path=db_path,
    )

    assert not has_orders_for_client(client.id or 0, db_path=db_path)
    archive_client(client.id or 0, db_path=db_path)
    assert list_clients(db_path=db_path) == []
    assert len(list_clients(include_archived=True, db_path=db_path)) == 1
