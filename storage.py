from __future__ import annotations

import csv
import hashlib
import os
import re
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Optional

from models import AppConfig, ClientRecord, LegacyClientSeed, SeedReport

APP_DATA_SUBDIR = "UneedTShirtsNewOrder"

HEADER_MAP = {
    "client name": "name",
    "contact person": "contact_person",
    "phone number": "phone",
    "client address": "address",
    "email": "email",
}

STATE_ALIASES = {
    "FLORIDA": "FL",
}

ADDRESS_PATTERN = re.compile(
    r"^(?P<street>.*?)(?:,\s*)?(?P<city>[A-Za-z0-9.\-\s']+?),\s*(?P<state>[A-Za-z.\s]+)\s*(?P<zip>\d{5}(?:-\d{4})?)?\s*$"
)

ABBR_TOKEN_RE = re.compile(r"[A-Z0-9]+")


def get_app_data_dir() -> Path:
    local_app_data = os.getenv("LOCALAPPDATA")
    if local_app_data:
        base = Path(local_app_data)
    else:
        base = Path.home() / ".local" / "share"

    target = base / APP_DATA_SUBDIR
    try:
        target.mkdir(parents=True, exist_ok=True)
        return target
    except OSError:
        temp_root = Path(os.getenv("TMPDIR", "/tmp"))
        fallback = temp_root / APP_DATA_SUBDIR
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback


def get_logs_dir() -> Path:
    logs_dir = get_app_data_dir() / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    return logs_dir


def get_db_path() -> Path:
    return get_app_data_dir() / "state.db"


def _connect(db_path: Optional[Path] = None) -> sqlite3.Connection:
    path = db_path or get_db_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def _table_has_column(conn: sqlite3.Connection, table_name: str, column_name: str) -> bool:
    rows = conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    return any(row["name"] == column_name for row in rows)


def ensure_schema(db_path: Optional[Path] = None) -> None:
    with _connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS clients (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL COLLATE NOCASE UNIQUE,
                abbr TEXT NOT NULL COLLATE NOCASE UNIQUE,
                folder_path TEXT NOT NULL,
                contact_person TEXT NOT NULL DEFAULT '',
                phone TEXT NOT NULL DEFAULT '',
                email TEXT NOT NULL DEFAULT '',
                street_address TEXT NOT NULL DEFAULT '',
                city_state_zip TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT 'active' CHECK(status IN ('active','archived')),
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS internal_order_ids (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                yymm TEXT NOT NULL,
                seq INTEGER NOT NULL,
                internal_order_id TEXT NOT NULL UNIQUE,
                created_at TEXT NOT NULL,
                client_id INTEGER,
                client_name TEXT,
                folder_job_number TEXT,
                folder_path TEXT,
                workbook_path TEXT,
                UNIQUE (yymm, seq)
            )
            """
        )

        if not _table_has_column(conn, "internal_order_ids", "client_id"):
            conn.execute("ALTER TABLE internal_order_ids ADD COLUMN client_id INTEGER")

        if not _table_has_column(conn, "clients", "template_id"):
            conn.execute("ALTER TABLE clients ADD COLUMN template_id TEXT NOT NULL DEFAULT ''")

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS app_meta (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
            """
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS client_job_descriptions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                client_id INTEGER NOT NULL,
                name TEXT NOT NULL,
                UNIQUE(client_id, name)
            )
            """
        )


def _normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip())


def _normalize_state(state_raw: str) -> str:
    state = _normalize_space(state_raw).replace(".", "").upper()
    if not state:
        return ""
    if state in STATE_ALIASES:
        return STATE_ALIASES[state]
    if len(state) == 2:
        return state
    return state


def parse_address_components(raw_address: str) -> tuple[str, str, bool]:
    address = _normalize_space(raw_address)
    if not address:
        return "", "", False

    match = ADDRESS_PATTERN.match(address)
    if not match:
        return address, "", True

    street = _normalize_space(match.group("street").rstrip(","))
    city = _normalize_space(match.group("city"))
    state = _normalize_state(match.group("state"))
    zip_code = _normalize_space(match.group("zip") or "")

    if not city or not state:
        return address, "", True

    city_state_zip = f"{city}, {state}"
    if zip_code:
        city_state_zip = f"{city_state_zip} {zip_code}"

    return street, city_state_zip, False


def generate_client_abbreviation(name: str, taken_abbreviations: set[str]) -> str:
    tokens = ABBR_TOKEN_RE.findall(name.upper())
    if not tokens:
        tokens = ["CLT"]

    base = "".join(token[0] for token in tokens[:4])
    if len(base) < 3:
        first = "".join(tokens)
        idx = 1
        while len(base) < 3 and idx < len(first):
            base += first[idx]
            idx += 1

    base = (base + "XXX")[: max(3, len(base))]
    candidate = base
    counter = 2
    while candidate in taken_abbreviations:
        candidate = f"{base}{counter}"
        counter += 1

    return candidate


def _row_to_client(row: sqlite3.Row) -> ClientRecord:
    keys = row.keys()
    return ClientRecord(
        id=row["id"],
        name=row["name"],
        abbr=row["abbr"],
        folder_path=row["folder_path"],
        contact_person=row["contact_person"],
        phone=row["phone"],
        email=row["email"],
        street_address=row["street_address"],
        city_state_zip=row["city_state_zip"],
        template_id=row["template_id"] if "template_id" in keys else "",
        status=row["status"],
        created_at=row["created_at"],
        updated_at=row["updated_at"],
    )


def list_clients(include_archived: bool = False, db_path: Optional[Path] = None) -> list[ClientRecord]:
    ensure_schema(db_path)
    with _connect(db_path) as conn:
        if include_archived:
            rows = conn.execute(
                "SELECT * FROM clients ORDER BY name COLLATE NOCASE ASC"
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM clients WHERE status='active' ORDER BY name COLLATE NOCASE ASC"
            ).fetchall()
    return [_row_to_client(row) for row in rows]


def get_client_by_id(client_id: int, db_path: Optional[Path] = None) -> Optional[ClientRecord]:
    ensure_schema(db_path)
    with _connect(db_path) as conn:
        row = conn.execute(
            "SELECT * FROM clients WHERE id = ?",
            (client_id,),
        ).fetchone()
    return _row_to_client(row) if row else None


def get_client_by_name(name: str, db_path: Optional[Path] = None) -> Optional[ClientRecord]:
    ensure_schema(db_path)
    with _connect(db_path) as conn:
        row = conn.execute(
            "SELECT * FROM clients WHERE name = ? COLLATE NOCASE",
            (name,),
        ).fetchone()
    return _row_to_client(row) if row else None


def _upsert_client_connection(conn: sqlite3.Connection, client: ClientRecord) -> ClientRecord:
    now = datetime.now().isoformat()
    name = _normalize_space(client.name)
    abbr = _normalize_space(client.abbr).upper()
    folder_path = _normalize_space(client.folder_path)
    status = client.status or "active"

    if client.id is None:
        existing = conn.execute(
            "SELECT id FROM clients WHERE name = ? COLLATE NOCASE",
            (name,),
        ).fetchone()
        client_id = int(existing["id"]) if existing else None
    else:
        client_id = client.id

    if client_id is None:
        cursor = conn.execute(
            """
            INSERT INTO clients (
                name, abbr, folder_path, contact_person, phone, email,
                street_address, city_state_zip, template_id, status, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                name,
                abbr,
                folder_path,
                client.contact_person,
                client.phone,
                client.email,
                client.street_address,
                client.city_state_zip,
                client.template_id,
                status,
                now,
                now,
            ),
        )
        client_id = int(cursor.lastrowid)
    else:
        conn.execute(
            """
            UPDATE clients
            SET name=?, abbr=?, folder_path=?, contact_person=?, phone=?, email=?,
                street_address=?, city_state_zip=?, template_id=?, status=?, updated_at=?
            WHERE id=?
            """,
            (
                name,
                abbr,
                folder_path,
                client.contact_person,
                client.phone,
                client.email,
                client.street_address,
                client.city_state_zip,
                client.template_id,
                status,
                now,
                client_id,
            ),
        )

    row = conn.execute("SELECT * FROM clients WHERE id=?", (client_id,)).fetchone()
    return _row_to_client(row)


def upsert_client(client: ClientRecord, db_path: Optional[Path] = None) -> ClientRecord:
    ensure_schema(db_path)
    with _connect(db_path) as conn:
        return _upsert_client_connection(conn, client)


def archive_client(client_id: int, db_path: Optional[Path] = None) -> None:
    ensure_schema(db_path)
    with _connect(db_path) as conn:
        conn.execute(
            "UPDATE clients SET status='archived', updated_at=? WHERE id=?",
            (datetime.now().isoformat(), client_id),
        )


def has_orders_for_client(client_id: int, db_path: Optional[Path] = None) -> bool:
    ensure_schema(db_path)
    with _connect(db_path) as conn:
        row = conn.execute(
            "SELECT COUNT(*) AS count FROM internal_order_ids WHERE client_id=?",
            (client_id,),
        ).fetchone()
    return bool(row and int(row["count"]) > 0)


def _set_meta(conn: sqlite3.Connection, key: str, value: str) -> None:
    conn.execute(
        "INSERT INTO app_meta(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
        (key, value),
    )


def _count_clients(conn: sqlite3.Connection) -> int:
    row = conn.execute("SELECT COUNT(*) AS count FROM clients").fetchone()
    return int(row["count"]) if row else 0


def _csv_sha256(csv_path: Path) -> str:
    digest = hashlib.sha256()
    digest.update(csv_path.read_bytes())
    return digest.hexdigest()


def _normalize_row(raw_row: dict[str, str]) -> dict[str, str]:
    normalized: dict[str, str] = {}
    for key, value in raw_row.items():
        normalized_key = _normalize_space((key or "").lower())
        mapped_key = HEADER_MAP.get(normalized_key)
        if mapped_key:
            normalized[mapped_key] = _normalize_space(value or "")
    return normalized


def import_clients_from_csv(
    config: AppConfig,
    csv_path: Path,
    db_path: Optional[Path] = None,
) -> SeedReport:
    ensure_schema(db_path)
    report = SeedReport()

    if not csv_path.exists():
        report.warnings.append(f"CSV file not found: {csv_path}")
        return report

    with _connect(db_path) as conn:
        existing_clients = conn.execute("SELECT * FROM clients").fetchall()
        by_name = {row["name"].lower(): _row_to_client(row) for row in existing_clients}
        taken_abbr = {row["abbr"].upper() for row in existing_clients if row["abbr"]}

        with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            for index, row in enumerate(reader, start=2):
                normalized = _normalize_row(row)
                name = normalized.get("name", "")
                if not name:
                    report.skipped += 1
                    report.warnings.append(f"Row {index}: missing Client name")
                    continue

                street_address, city_state_zip, warning = parse_address_components(
                    normalized.get("address", "")
                )
                if warning and normalized.get("address", ""):
                    report.warnings.append(
                        f"Row {index} ({name}): could not confidently split address"
                    )

                existing = by_name.get(name.lower())
                if existing:
                    abbr = existing.abbr.strip().upper()
                    if not abbr:
                        abbr = generate_client_abbreviation(name, taken_abbr)
                        taken_abbr.add(abbr)

                    updated_record = ClientRecord(
                        id=existing.id,
                        name=existing.name,
                        abbr=abbr,
                        folder_path=existing.folder_path or str(Path(config.root_paths.clients_root) / name),
                        contact_person=normalized.get("contact_person", ""),
                        phone=normalized.get("phone", ""),
                        email=normalized.get("email", ""),
                        street_address=street_address,
                        city_state_zip=city_state_zip,
                        status=existing.status,
                        created_at=existing.created_at,
                        updated_at=existing.updated_at,
                    )
                    persisted = _upsert_client_connection(conn, updated_record)
                    by_name[persisted.name.lower()] = persisted
                    report.updated += 1
                else:
                    abbr = generate_client_abbreviation(name, taken_abbr)
                    taken_abbr.add(abbr)
                    new_record = ClientRecord(
                        name=name,
                        abbr=abbr,
                        folder_path=str(Path(config.root_paths.clients_root) / name),
                        contact_person=normalized.get("contact_person", ""),
                        phone=normalized.get("phone", ""),
                        email=normalized.get("email", ""),
                        street_address=street_address,
                        city_state_zip=city_state_zip,
                        status="active",
                    )
                    persisted = _upsert_client_connection(conn, new_record)
                    by_name[persisted.name.lower()] = persisted
                    report.created += 1

        _set_meta(conn, "clients_csv_sha256", _csv_sha256(csv_path))
        _set_meta(conn, "clients_csv_imported_at", datetime.now().isoformat())

    return report


def _seed_clients_from_legacy_config(
    config: AppConfig,
    db_path: Optional[Path] = None,
) -> SeedReport:
    ensure_schema(db_path)
    report = SeedReport()
    if not config.legacy_clients:
        return report

    with _connect(db_path) as conn:
        existing_clients = conn.execute("SELECT * FROM clients").fetchall()
        by_name = {row["name"].lower(): _row_to_client(row) for row in existing_clients}
        taken_abbr = {row["abbr"].upper() for row in existing_clients if row["abbr"]}

        for legacy in config.legacy_clients:
            existing = by_name.get(legacy.name.lower())
            abbr = (legacy.abbr or "").upper()
            if not abbr:
                abbr = generate_client_abbreviation(legacy.name, taken_abbr)
            if abbr in taken_abbr and (existing is None or existing.abbr.upper() != abbr):
                abbr = generate_client_abbreviation(legacy.name, taken_abbr)

            taken_abbr.add(abbr)
            folder_path = str(Path(config.root_paths.clients_root) / legacy.folder_name)

            record = ClientRecord(
                id=existing.id if existing else None,
                name=legacy.name,
                abbr=abbr,
                folder_path=existing.folder_path if existing and existing.folder_path else folder_path,
                contact_person=legacy.contact_person,
                phone=legacy.phone,
                email=legacy.email,
                street_address=legacy.street_address,
                city_state_zip=legacy.city_state_zip,
                status=(existing.status if existing else "active"),
                created_at=(existing.created_at if existing else ""),
                updated_at=(existing.updated_at if existing else ""),
            )
            persisted = _upsert_client_connection(conn, record)
            by_name[persisted.name.lower()] = persisted
            if existing:
                report.updated += 1
            else:
                report.created += 1

    return report


def seed_clients_from_csv_if_needed(
    config: AppConfig,
    csv_path: Path,
    db_path: Optional[Path] = None,
) -> SeedReport:
    ensure_schema(db_path)
    with _connect(db_path) as conn:
        if _count_clients(conn) > 0:
            return SeedReport()

    if csv_path.exists():
        return import_clients_from_csv(config, csv_path, db_path=db_path)

    return _seed_clients_from_legacy_config(config, db_path=db_path)


def next_internal_order_id(now: datetime, db_path: Optional[Path] = None, _max_retries: int = 3) -> str:
    ensure_schema(db_path)
    yymm = now.strftime("%y%m")
    for attempt in range(_max_retries):
        try:
            with _connect(db_path) as conn:
                current = conn.execute(
                    "SELECT COALESCE(MAX(seq), 0) AS max_seq FROM internal_order_ids WHERE yymm = ?",
                    (yymm,),
                ).fetchone()
                next_seq = (int(current["max_seq"]) if current else 0) + 1
                internal_order_id = f"{yymm}-{next_seq:04d}"
                conn.execute(
                    """
                    INSERT INTO internal_order_ids (yymm, seq, internal_order_id, created_at)
                    VALUES (?, ?, ?, ?)
                    """,
                    (yymm, next_seq, internal_order_id, now.isoformat()),
                )
            return internal_order_id
        except sqlite3.IntegrityError:
            if attempt == _max_retries - 1:
                raise
    raise sqlite3.IntegrityError("Failed to generate unique internal order ID")


def record_order_event(
    internal_order_id: str,
    client_name: str,
    folder_job_number: str,
    folder_path: str,
    workbook_path: str,
    client_id: Optional[int] = None,
    db_path: Optional[Path] = None,
) -> None:
    ensure_schema(db_path)
    try:
        yymm, seq_text = internal_order_id.split("-")
        seq = int(seq_text)
    except ValueError:
        yymm = ""
        seq = 0

    now = datetime.now().isoformat()
    with _connect(db_path) as conn:
        cursor = conn.execute(
            """
            UPDATE internal_order_ids
            SET client_id = ?, client_name = ?, folder_job_number = ?, folder_path = ?, workbook_path = ?
            WHERE internal_order_id = ?
            """,
            (client_id, client_name, folder_job_number, folder_path, workbook_path, internal_order_id),
        )
        if cursor.rowcount == 0:
            conn.execute(
                """
                INSERT INTO internal_order_ids (
                    yymm,
                    seq,
                    internal_order_id,
                    created_at,
                    client_id,
                    client_name,
                    folder_job_number,
                    folder_path,
                    workbook_path
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    yymm,
                    seq,
                    internal_order_id,
                    now,
                    client_id,
                    client_name,
                    folder_job_number,
                    folder_path,
                    workbook_path,
                ),
            )


def list_job_descriptions(client_id: int, db_path: Optional[Path] = None) -> list[str]:
    ensure_schema(db_path)
    with _connect(db_path) as conn:
        rows = conn.execute(
            "SELECT name FROM client_job_descriptions WHERE client_id=? ORDER BY name COLLATE NOCASE ASC",
            (client_id,),
        ).fetchall()
    return [row["name"] for row in rows]


def list_job_description_records(client_id: int, db_path: Optional[Path] = None) -> list[tuple[int, str]]:
    ensure_schema(db_path)
    with _connect(db_path) as conn:
        rows = conn.execute(
            "SELECT id, name FROM client_job_descriptions WHERE client_id=? ORDER BY name COLLATE NOCASE ASC",
            (client_id,),
        ).fetchall()
    return [(row["id"], row["name"]) for row in rows]


def add_job_description(client_id: int, name: str, db_path: Optional[Path] = None) -> None:
    ensure_schema(db_path)
    with _connect(db_path) as conn:
        conn.execute(
            "INSERT INTO client_job_descriptions (client_id, name) VALUES (?, ?)",
            (client_id, name.strip()),
        )


def update_job_description(desc_id: int, name: str, db_path: Optional[Path] = None) -> None:
    ensure_schema(db_path)
    with _connect(db_path) as conn:
        conn.execute(
            "UPDATE client_job_descriptions SET name=? WHERE id=?",
            (name.strip(), desc_id),
        )


def delete_job_description(desc_id: int, db_path: Optional[Path] = None) -> None:
    ensure_schema(db_path)
    with _connect(db_path) as conn:
        conn.execute(
            "DELETE FROM client_job_descriptions WHERE id=?",
            (desc_id,),
        )
