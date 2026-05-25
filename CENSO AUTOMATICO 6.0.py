from __future__ import annotations

import importlib.util
import logging
import math
import os
import re
import subprocess
import sys
import unicodedata
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any


BASE_DIR = Path(__file__).resolve().parent
REQUIREMENTS_PATH = BASE_DIR / "requirements.txt"
REQUIRED_IMPORTS = {
    "bs4": "beautifulsoup4",
    "gspread": "gspread",
    "oauth2client": "oauth2client",
    "openpyxl": "openpyxl",
    "pandas": "pandas",
    "requests": "requests",
    "xlrd": "xlrd",
}


def ensure_dependencies() -> None:
    missing = [
        package_name
        for import_name, package_name in REQUIRED_IMPORTS.items()
        if importlib.util.find_spec(import_name) is None
    ]
    if not missing:
        return

    if not REQUIREMENTS_PATH.exists():
        missing_text = ", ".join(missing)
        raise RuntimeError(
            f"Faltan librerias: {missing_text}. No se encontro requirements.txt para instalarlas."
        )

    print("Faltan librerias necesarias. Se instalaran automaticamente.")
    print("Esto puede demorar unos minutos la primera vez.")
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "-r", str(REQUIREMENTS_PATH)]
    )

    still_missing = [
        package_name
        for import_name, package_name in REQUIRED_IMPORTS.items()
        if importlib.util.find_spec(import_name) is None
    ]
    if still_missing:
        missing_text = ", ".join(still_missing)
        raise RuntimeError(f"No se pudieron instalar estas librerias: {missing_text}")


ensure_dependencies()

import gspread
import pandas as pd
import requests
from bs4 import BeautifulSoup
from gspread.cell import Cell
from oauth2client.service_account import ServiceAccountCredentials

DEFAULT_SPREADSHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1E49TZkkm_d9qbyAHfDyZ_zDoqmjzdOG6IGAqAegfW3k/edit?gid=0#gid=0"
)
DEFAULT_LOGIN_URL = "http://10.6.84.181/login"
DEFAULT_PATIENT_LIST_URL_TEMPLATE = (
    "http://10.6.84.191/proyectos/gestion_clinica/index.php?id_sector={sector_id}"
)
DEFAULT_EXCEL_URL_TEMPLATE = (
    "http://10.6.84.191/proyectos/gestion_clinica/ListaSectorExcel.php?fse_id={sector_id}"
)
DEFAULT_TARGET_SHEET_NAMES = [
    "S. DE NEUROCIRUGIA",
    "UCI/UTI",
    "PEDIATRICOS",
    "OTROS SERVICIOS",
]
DEFAULT_SOURCE_SECTORS = [
    {"sector_id": 166, "target_sheet": "S. DE NEUROCIRUGIA", "label": "Servicio Neurocirugia"},
    {"sector_id": 15, "target_sheet": "UCI/UTI", "label": "UCI 2do piso"},
    {"sector_id": 5, "target_sheet": "UCI/UTI", "label": "UTI 2do piso"},
    {"sector_id": 6, "target_sheet": "UCI/UTI", "label": "UCI/UTI 3ro"},
    {"sector_id": 16, "target_sheet": "PEDIATRICOS", "label": "UCI/UTI Pediatricos"},
    {
        "sector_id": 14,
        "target_sheet": "OTROS SERVICIOS",
        "label": "Urgencia",
        "filter_column": "Especialidad",
        "filter_value": "NEUROCIRUGIA",
    },
    {
        "sector_id": 2,
        "target_sheet": "OTROS SERVICIOS",
        "label": "TMT",
        "filter_column": "Especialidad",
        "filter_value": "NEUROCIRUGIA",
    },
]
DEFAULT_CREDENTIAL_FILENAMES = [
    "service_account.json",
    "credentials.json",
    "censo-automatico-430916-f23cd66e6932.json",
]

GOOGLE_SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]
AUTO_FIELD_KEYS = ["patient_name", "diagnosis", "rut"]
PRESERVED_FIELD_KEYS = ["staff", "becado", "plan"]
STAFF_BECADO_FIELD_KEYS = ["staff", "becado"]
REQUIRED_SCHEMA_KEYS = ["bed", "patient_name", "diagnosis", "rut"]

HEADER_ALIASES = {
    "bed": ["CAMA", "SALA", "HABITACION", "BOX"],
    "patient_name": ["NOMBRE", "PACIENTE", "NOMBRE PACIENTE"],
    "diagnosis": ["DIAGNOSTICO", "DIAG", "DX"],
    "rut": ["RUT", "RUN", "ID", "IDENTIFICACION"],
    "staff": ["STAFF"],
    "becado": ["BECADO", "BECADA", "RESIDENTE", "BECARIO"],
    "plan": ["PLAN", "OBSERVACIONES", "OBSERVACION", "OBS"],
}


@dataclass(frozen=True)
class SectorConfig:
    sector_id: int
    target_sheet: str
    label: str
    filter_column: str | None = None
    filter_value: str | None = None


@dataclass(frozen=True)
class AppConfig:
    spreadsheet_url: str
    login_url: str
    patient_list_url_template: str
    excel_url_template: str
    username: str
    password: str
    credentials_path: Path
    target_sheet_names: list[str]
    source_sectors: list[SectorConfig]
    clear_missing_beds: bool = False


def setup_logging() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


def load_dotenv(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.exists():
        return values

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def load_project_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}

    spec = importlib.util.spec_from_file_location("censo_project_config", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load config file: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return {
        key: getattr(module, key)
        for key in dir(module)
        if key.isupper() and not key.startswith("_")
    }


def first_config_value(
    key: str,
    env: dict[str, str],
    dotenv_values: dict[str, str],
    project_config: dict[str, Any],
    default: Any = None,
) -> Any:
    if key in env and env[key] != "":
        return env[key]
    if key in dotenv_values and dotenv_values[key] != "":
        return dotenv_values[key]
    return project_config.get(key, default)


def required_config_value(
    key: str,
    env: dict[str, str],
    dotenv_values: dict[str, str],
    project_config: dict[str, Any],
) -> Any:
    value = first_config_value(key, env, dotenv_values, project_config)
    if value in (None, ""):
        raise RuntimeError(
            f"Missing required configuration value {key}. "
            "Set it in the environment, .env, or config.py."
        )
    return value


def resolve_path(value: str | Path | None) -> Path | None:
    if not value:
        return None
    path = Path(value).expanduser()
    if path.is_absolute():
        return path
    return BASE_DIR / path


def resolve_credentials_path(
    env: dict[str, str],
    dotenv_values: dict[str, str],
    project_config: dict[str, Any],
) -> Path:
    candidates = [
        resolve_path(env.get("CENSO_SERVICE_ACCOUNT_JSON")),
        resolve_path(env.get("GOOGLE_APPLICATION_CREDENTIALS")),
        resolve_path(dotenv_values.get("CENSO_SERVICE_ACCOUNT_JSON")),
        resolve_path(project_config.get("CREDENTIALS_PATH")),
    ]
    candidates.extend(BASE_DIR / filename for filename in DEFAULT_CREDENTIAL_FILENAMES)

    for path in candidates:
        if path and path.exists():
            return path

    checked = ", ".join(str(path) for path in candidates if path)
    raise FileNotFoundError(
        "Google service-account credentials were not found. "
        "Set CENSO_SERVICE_ACCOUNT_JSON, add it to .env, set CREDENTIALS_PATH in config.py, "
        f"or place a credentials file beside this script. Checked: {checked}"
    )


def load_config() -> AppConfig:
    dotenv_values = load_dotenv(BASE_DIR / ".env")
    project_config = load_project_config(BASE_DIR / "config.py")
    env = dict(os.environ)

    source_sector_values = first_config_value(
        "SOURCE_SECTORS", env, dotenv_values, project_config, DEFAULT_SOURCE_SECTORS
    )
    source_sectors = [SectorConfig(**sector) for sector in source_sector_values]

    username = required_config_value("CENSO_USERNAME", env, dotenv_values, project_config)
    password = required_config_value("CENSO_PASSWORD", env, dotenv_values, project_config)

    return AppConfig(
        spreadsheet_url=first_config_value(
            "SPREADSHEET_URL", env, dotenv_values, project_config, DEFAULT_SPREADSHEET_URL
        ),
        login_url=first_config_value("LOGIN_URL", env, dotenv_values, project_config, DEFAULT_LOGIN_URL),
        patient_list_url_template=first_config_value(
            "PATIENT_LIST_URL_TEMPLATE",
            env,
            dotenv_values,
            project_config,
            DEFAULT_PATIENT_LIST_URL_TEMPLATE,
        ),
        excel_url_template=first_config_value(
            "EXCEL_URL_TEMPLATE", env, dotenv_values, project_config, DEFAULT_EXCEL_URL_TEMPLATE
        ),
        username=username,
        password=password,
        credentials_path=resolve_credentials_path(env, dotenv_values, project_config),
        target_sheet_names=list(
            first_config_value("TARGET_SHEET_NAMES", env, dotenv_values, project_config, DEFAULT_TARGET_SHEET_NAMES)
        ),
        source_sectors=source_sectors,
        clear_missing_beds=str(
            first_config_value("CLEAR_MISSING_BEDS", env, dotenv_values, project_config, "false")
        ).lower()
        in {"1", "true", "yes", "y"},
    )


def clean_cell_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat"}:
        return ""
    return text


def normalize_text(value: Any, keep_rut_chars: bool = False) -> str:
    text = clean_cell_value(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = text.upper().strip()
    if keep_rut_chars:
        text = re.sub(r"[^0-9K]", "", text)
    else:
        text = re.sub(r"[^A-Z0-9 ]+", " ", text)
        text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_rut(value: Any) -> str:
    rut = normalize_text(value, keep_rut_chars=True)
    return rut if len(rut) >= 5 else ""


def normalize_name(value: Any) -> str:
    return normalize_text(value)


def normalize_header(value: Any) -> str:
    return normalize_text(value)


def normalize_bed(value: Any) -> str:
    text = normalize_text(value)
    text = re.sub(r"\b(CAMA|SALA|HABITACION|HAB|BOX)\b", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    tokens = [str(int(token)) if token.isdigit() else token for token in text.split()]
    return " ".join(tokens)


def find_column_index(row: pd.Series, aliases: list[str]) -> str | None:
    normalized_aliases = {normalize_header(alias) for alias in aliases}
    for column in row.index:
        if normalize_header(column) in normalized_aliases:
            return str(column)
    return None


def get_first_value(row: pd.Series, aliases: list[str]) -> str:
    column = find_column_index(row, aliases)
    if column is None:
        return ""
    return clean_cell_value(row.get(column, ""))


def create_session() -> requests.Session:
    return requests.Session()


def login_to_hospital_system(session: requests.Session, username: str, password: str, login_url: str) -> None:
    response = session.get(login_url, timeout=30)
    response.raise_for_status()

    soup = BeautifulSoup(response.content, "html.parser")
    csrf_input = soup.find("input", {"name": "_token"})
    if csrf_input is None or not csrf_input.get("value"):
        raise RuntimeError("Could not find login CSRF token.")

    payload = {
        "usu_login": username,
        "usu_password": password,
        "_token": csrf_input["value"],
    }
    login_response = session.post(login_url, data=payload, timeout=30)
    login_response.raise_for_status()
    logging.info("Login successful")


def download_sector_dataframe(
    session: requests.Session,
    sector_id: int,
    excel_url_template: str,
    patient_list_url_template: str | None = None,
    header_row: int = 5,
) -> pd.DataFrame:
    if patient_list_url_template:
        patient_list_url = patient_list_url_template.format(sector_id=sector_id)
        navigation_response = session.get(patient_list_url, timeout=30)
        navigation_response.raise_for_status()

    excel_url = excel_url_template.format(sector_id=sector_id)
    response = session.get(excel_url, timeout=30)
    response.raise_for_status()
    try:
        return pd.read_excel(BytesIO(response.content), header=header_row)
    except ValueError as exc:
        if "Passed header" in str(exc) and "only" in str(exc) and "lines in file" in str(exc):
            logging.warning(
                "Sector %s export did not contain row %s headers; treating it as empty. Original error: %s",
                sector_id,
                header_row + 1,
                exc,
            )
            return pd.DataFrame()
        raise


def authorize_google_sheets(credentials_path: Path) -> gspread.Client:
    credentials = ServiceAccountCredentials.from_json_keyfile_name(str(credentials_path), GOOGLE_SCOPE)
    return gspread.authorize(credentials)


def get_target_worksheets(spreadsheet: gspread.Spreadsheet, target_sheet_names: list[str]) -> dict[str, gspread.Worksheet]:
    first_four = spreadsheet.worksheets()[:4]
    first_four_by_title = {worksheet.title: worksheet for worksheet in first_four}
    missing = [name for name in target_sheet_names if name not in first_four_by_title]
    if missing:
        first_four_titles = [worksheet.title for worksheet in first_four]
        raise RuntimeError(
            "Target sheets must be present in the first 4 workbook tabs. "
            f"Missing: {missing}. First 4 tabs: {first_four_titles}"
        )
    return {name: first_four_by_title[name] for name in target_sheet_names}


def get_worksheet_schema(worksheet: gspread.Worksheet) -> dict[str, Any]:
    values = worksheet.get_all_values()
    best: tuple[int, int, dict[str, int], dict[str, list[int]]] | None = None

    for row_index, row in enumerate(values[:10], start=1):
        column_map: dict[str, int] = {}
        multi_column_map: dict[str, list[int]] = {}
        normalized_cells = [normalize_header(value) for value in row]
        for key, aliases in HEADER_ALIASES.items():
            matched_columns: list[int] = []
            for alias in aliases:
                normalized_alias = normalize_header(alias)
                for column_index, normalized_cell in enumerate(normalized_cells, start=1):
                    if normalized_cell == normalized_alias and column_index not in matched_columns:
                        matched_columns.append(column_index)
                        break
            if matched_columns:
                column_map[key] = matched_columns[0]
                multi_column_map[key] = matched_columns

        score = sum(1 for key in REQUIRED_SCHEMA_KEYS if key in column_map)
        if best is None or score > best[0]:
            best = (score, row_index, column_map, multi_column_map)

    if best is None or best[0] < 3 or "bed" not in best[2]:
        raise RuntimeError(
            f"Could not detect a usable header row in worksheet {worksheet.title}. "
            "Expected at least bed, patient, diagnosis/RUT style headers."
        )

    missing_auto = [key for key in REQUIRED_SCHEMA_KEYS if key not in best[2]]
    if missing_auto:
        raise RuntimeError(f"Worksheet {worksheet.title} is missing required columns: {missing_auto}")

    return {
        "header_row": best[1],
        "columns": best[2],
        "multi_columns": best[3],
        "values": values,
    }


def get_cell(values: list[list[str]], row_number: int, column_number: int) -> str:
    row_index = row_number - 1
    column_index = column_number - 1
    if row_index < 0 or row_index >= len(values):
        return ""
    row = values[row_index]
    if column_index < 0 or column_index >= len(row):
        return ""
    return clean_cell_value(row[column_index])


def get_bed_value_from_worksheet_row(
    values: list[list[str]],
    row_number: int,
    schema: dict[str, Any],
    use_key_columns: bool = False,
) -> str:
    bed_columns = (
        schema.get("bed_key_columns")
        if use_key_columns
        else schema.get("multi_columns", {}).get("bed")
    ) or [schema["columns"]["bed"]]
    parts = [get_cell(values, row_number, column_number) for column_number in bed_columns]
    return " ".join(part for part in parts if part)


def build_bed_row_map(
    worksheet: gspread.Worksheet,
    schema: dict[str, Any],
    bed_column_name: str = "bed",
) -> dict[str, int]:
    values = schema["values"]
    header_row = schema["header_row"]
    primary_bed_column = schema["columns"][bed_column_name]
    all_bed_columns = schema.get("multi_columns", {}).get("bed") or [primary_bed_column]
    primary_counts: dict[str, int] = {}

    for row_number in range(header_row + 1, len(values) + 1):
        primary_key = normalize_bed(get_cell(values, row_number, primary_bed_column))
        if primary_key:
            primary_counts[primary_key] = primary_counts.get(primary_key, 0) + 1

    has_duplicate_primary_beds = any(count > 1 for count in primary_counts.values())
    if has_duplicate_primary_beds:
        bed_key_columns = list(all_bed_columns)
        group_column = primary_bed_column - 1
        if group_column >= 1 and group_column not in bed_key_columns:
            group_values = {
                normalize_bed(get_cell(values, row_number, group_column))
                for row_number in range(header_row + 1, len(values) + 1)
            }
            group_values.discard("")
            if group_values:
                bed_key_columns.append(group_column)
                logging.info(
                    "Worksheet %s has repeated bed values; using columns %s as the bed identity",
                    worksheet.title,
                    bed_key_columns,
                )
        schema["bed_key_columns"] = bed_key_columns
    else:
        schema["bed_key_columns"] = [primary_bed_column]

    bed_map: dict[str, int] = {}

    for row_number in range(header_row + 1, len(values) + 1):
        primary_key = normalize_bed(get_cell(values, row_number, primary_bed_column))
        if not primary_key:
            continue
        bed_value = get_bed_value_from_worksheet_row(values, row_number, schema, use_key_columns=True)
        bed_key = normalize_bed(bed_value)
        if not bed_key:
            continue
        if bed_key in bed_map:
            logging.warning(
                "Duplicate bed %s in worksheet %s rows %s and %s; using first row",
                bed_key,
                worksheet.title,
                bed_map[bed_key],
                row_number,
            )
            continue
        bed_map[bed_key] = row_number
    return bed_map


def build_existing_patient_snapshot(
    worksheets: dict[str, gspread.Worksheet],
    target_sheet_names: list[str],
) -> dict[str, dict[str, dict[str, str]]]:
    snapshot: dict[str, dict[str, dict[str, str]]] = {"rut": {}, "name": {}}

    for sheet_name in target_sheet_names:
        worksheet = worksheets[sheet_name]
        schema = get_worksheet_schema(worksheet)
        values = schema["values"]
        columns = schema["columns"]
        logging.info("Reading existing patients from %s", sheet_name)

        for row_number in range(schema["header_row"] + 1, len(values) + 1):
            record = {
                "patient_name": get_cell(values, row_number, columns["patient_name"]),
                "rut": get_cell(values, row_number, columns["rut"]),
                "staff": get_cell(values, row_number, columns.get("staff", 0)),
                "becado": get_cell(values, row_number, columns.get("becado", 0)),
                "plan": get_cell(values, row_number, columns.get("plan", 0)),
            }
            if not record["patient_name"] and not record["rut"]:
                continue

            preserved = {
                "staff": record["staff"],
                "becado": record["becado"],
                "plan": record["plan"],
                "source_sheet": sheet_name,
                "source_row": str(row_number),
                "source_bed": get_bed_value_from_worksheet_row(values, row_number, schema),
                "patient_name": record["patient_name"],
                "rut": record["rut"],
            }

            rut_key = normalize_rut(record["rut"])
            name_key = normalize_name(record["patient_name"])
            if rut_key and rut_key not in snapshot["rut"]:
                snapshot["rut"][rut_key] = preserved
            if name_key and name_key not in snapshot["name"]:
                snapshot["name"][name_key] = preserved

    logging.info(
        "Existing snapshot: %s RUT identities, %s name fallback identities",
        len(snapshot["rut"]),
        len(snapshot["name"]),
    )
    return snapshot


def apply_sector_filter(df: pd.DataFrame, sector: SectorConfig) -> pd.DataFrame:
    if not sector.filter_column or sector.filter_value is None:
        return df

    column = find_column_index(pd.Series(index=df.columns, dtype=object), [sector.filter_column])
    if column is None:
        logging.warning(
            "Sector %s requested filter column %s, but the column is not present",
            sector.label,
            sector.filter_column,
        )
        return df.iloc[0:0]

    expected = normalize_text(sector.filter_value)
    return df[df[column].map(normalize_text) == expected].reset_index(drop=True)


def get_bed_values_from_source_row(
    row: pd.Series,
    preferred_bed_header_groups: list[list[str]] | None = None,
) -> list[str]:
    if preferred_bed_header_groups:
        values = [get_first_value(row, headers) for headers in preferred_bed_header_groups if headers]
        values = [value for value in values if value]
        if values:
            return values

    cama = get_first_value(row, ["Cama"])
    sala = get_first_value(row, ["Sala"])
    if cama:
        return [cama]
    if sala:
        return [sala]
    fallback = get_first_value(row, HEADER_ALIASES["bed"])
    return [fallback] if fallback else []


def make_bed_value(row: pd.Series, preferred_bed_header_groups: list[list[str]] | None = None) -> str:
    return " ".join(get_bed_values_from_source_row(row, preferred_bed_header_groups))


def infer_target_group_value(row: pd.Series, sector: SectorConfig) -> str:
    if sector.target_sheet == "S. DE NEUROCIRUGIA":
        sala = get_first_value(row, ["Sala"])
        match = re.search(r"\d+", sala)
        return match.group(0) if match else sala

    if sector.target_sheet == "UCI/UTI":
        if sector.sector_id == 15:
            return "UCI 2"
        if sector.sector_id == 5:
            return "UTI 2"
        if sector.sector_id == 6:
            service_text = normalize_text(
                " ".join(
                    [
                        get_first_value(row, ["Servicio"]),
                        get_first_value(row, ["Especialidad"]),
                        get_first_value(row, ["Sala"]),
                    ]
                )
            )
            return "UTI 3" if "UTI" in service_text else "UCI 3"

    if sector.target_sheet == "PEDIATRICOS":
        service_text = normalize_text(
            " ".join(
                [
                    get_first_value(row, ["Servicio"]),
                    get_first_value(row, ["Especialidad"]),
                    get_first_value(row, ["Sala"]),
                ]
            )
        )
        return "UTI" if "UTI" in service_text else "UCI"

    return ""


def transform_sector_data_into_target_rows(
    df: pd.DataFrame,
    sector: SectorConfig,
    destination_schema: dict[str, Any],
) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    bed_key_columns = destination_schema.get("bed_key_columns") or [destination_schema["columns"]["bed"]]
    preferred_bed_header_groups: list[list[str]] = []
    primary_bed_column = destination_schema["columns"]["bed"]
    for column_number in bed_key_columns:
        header = get_cell(destination_schema["values"], destination_schema["header_row"], column_number)
        if header:
            preferred_bed_header_groups.append([header])
        elif column_number == primary_bed_column - 1:
            preferred_bed_header_groups.append(["Sala", "Habitacion", "Habitación", "Unidad"])
        else:
            preferred_bed_header_groups.append(["Cama"])

    for _, row in df.iterrows():
        patient_name = get_first_value(row, HEADER_ALIASES["patient_name"])
        rut = get_first_value(row, HEADER_ALIASES["rut"])
        diagnosis = get_first_value(row, HEADER_ALIASES["diagnosis"])
        service_text = normalize_text(get_first_value(row, ["Servicio"]))
        if not patient_name and not rut:
            continue
        if service_text == "DESOCUPADA":
            continue

        bed_values = get_bed_values_from_source_row(row, preferred_bed_header_groups)
        target_group = infer_target_group_value(row, sector)
        if target_group and len(bed_values) > 1:
            bed_values[1] = target_group
        bed = " ".join(bed_values)

        if not bed and not patient_name and not rut:
            continue

        records.append(
            {
                "target_sheet": sector.target_sheet,
                "source_sector": str(sector.sector_id),
                "source_label": sector.label,
                "bed": bed,
                "primary_bed": bed_values[0] if bed_values else "",
                "bed_key": normalize_bed(bed),
                "patient_name": patient_name,
                "diagnosis": diagnosis,
                "rut": rut,
            }
        )

    return records


def resolve_preserved_fields(
    incoming: dict[str, str],
    snapshot: dict[str, dict[str, dict[str, str]]],
) -> tuple[str | None, dict[str, str] | None]:
    rut_key = normalize_rut(incoming.get("rut", ""))
    if rut_key and rut_key in snapshot["rut"]:
        return "rut", snapshot["rut"][rut_key]

    name_key = normalize_name(incoming.get("patient_name", ""))
    if name_key and name_key in snapshot["name"]:
        return "name", snapshot["name"][name_key]

    return None, None


def set_cell_if_changed(
    updates: list[Cell],
    values: list[list[str]],
    row_number: int,
    column_number: int | None,
    value: str,
) -> bool:
    if not column_number:
        return False
    value = clean_cell_value(value)
    if get_cell(values, row_number, column_number) == value:
        return False
    updates.append(Cell(row_number, column_number, value))
    return True


def update_target_worksheet(
    worksheet: gspread.Worksheet,
    schema: dict[str, Any],
    bed_row_map: dict[str, int],
    incoming_records: list[dict[str, str]],
    snapshot: dict[str, dict[str, dict[str, str]]],
    clear_missing_beds: bool = False,
    moved_manual_clear_requests: list[tuple[str, int]] | None = None,
    protected_manual_rows: set[tuple[str, int]] | None = None,
) -> dict[str, int]:
    logging.info("Processing worksheet %s", worksheet.title)
    values = schema["values"]
    columns = schema["columns"]
    updates: list[Cell] = []
    incoming_beds: set[str] = set()
    stats = {
        "source_patients": len(incoming_records),
        "updated": 0,
        "missing_bed": 0,
        "matched_by_rut": 0,
        "matched_by_name": 0,
        "cleared_beds": 0,
    }

    for incoming in incoming_records:
        bed_key = incoming["bed_key"]
        match_method, preserved = resolve_preserved_fields(incoming, snapshot)
        if not bed_key or bed_key not in bed_row_map:
            if preserved and preserved.get("source_sheet") == worksheet.title and preserved.get("source_row", "").isdigit():
                row_number = int(preserved["source_row"])
                logging.info(
                    "Using existing patient row %s in worksheet %s while filling missing bed identity %r for %s",
                    row_number,
                    worksheet.title,
                    incoming.get("bed", ""),
                    incoming.get("patient_name", ""),
                )
            else:
                logging.warning(
                    "No destination row for bed %r in worksheet %s: %s",
                    incoming.get("bed", ""),
                    worksheet.title,
                    incoming.get("patient_name", ""),
                )
                stats["missing_bed"] += 1
                continue
        else:
            row_number = bed_row_map[bed_key]

        target_location = (worksheet.title, row_number)

        incoming_beds.add(bed_key)
        if match_method == "rut":
            stats["matched_by_rut"] += 1
        elif match_method == "name":
            stats["matched_by_name"] += 1

        changed = False
        if incoming.get("primary_bed") and not get_cell(values, row_number, columns["bed"]):
            changed |= set_cell_if_changed(updates, values, row_number, columns["bed"], incoming["primary_bed"])
        changed |= set_cell_if_changed(updates, values, row_number, columns["patient_name"], incoming["patient_name"])
        changed |= set_cell_if_changed(updates, values, row_number, columns["diagnosis"], incoming["diagnosis"])
        changed |= set_cell_if_changed(updates, values, row_number, columns["rut"], incoming["rut"])

        if preserved:
            if protected_manual_rows is not None:
                protected_manual_rows.add(target_location)

            for field in PRESERVED_FIELD_KEYS:
                changed |= set_cell_if_changed(updates, values, row_number, columns.get(field), preserved.get(field, ""))

            source_sheet = preserved.get("source_sheet", worksheet.title)
            source_row = preserved.get("source_row", "")
            source_bed = normalize_bed(preserved.get("source_bed", ""))
            source_location = (
                source_sheet,
                int(source_row),
            ) if source_row.isdigit() else None
            patient_moved = source_location is not None and source_location != target_location

            if patient_moved and moved_manual_clear_requests is not None:
                moved_manual_clear_requests.append(source_location)

            if source_sheet != worksheet.title or source_bed != bed_key:
                logging.info(
                    "Preserved STAFF/BECADO/PLAN for moved patient %s from %s bed %s to %s bed %s",
                    incoming.get("patient_name", ""),
                    source_sheet,
                    preserved.get("source_bed", ""),
                    worksheet.title,
                    incoming.get("bed", ""),
                )
        else:
            for field in STAFF_BECADO_FIELD_KEYS:
                changed |= set_cell_if_changed(updates, values, row_number, columns.get(field), "")

        if changed:
            stats["updated"] += 1

    if clear_missing_beds:
        occupied_existing_beds = set()
        for bed_key, row_number in bed_row_map.items():
            existing_name = get_cell(values, row_number, columns["patient_name"])
            existing_rut = get_cell(values, row_number, columns["rut"])
            if existing_name or existing_rut:
                occupied_existing_beds.add(bed_key)

        for bed_key in sorted(occupied_existing_beds - incoming_beds):
            row_number = bed_row_map[bed_key]
            cleared = False
            for field in AUTO_FIELD_KEYS:
                cleared |= set_cell_if_changed(updates, values, row_number, columns[field], "")
            if cleared:
                stats["cleared_beds"] += 1
                logging.info(
                    "Cleared empty bed %s in worksheet %s",
                    get_bed_value_from_worksheet_row(values, row_number, schema),
                    worksheet.title,
                )
    else:
        logging.info("Skipping empty-bed clearing in %s; CLEAR_MISSING_BEDS is disabled", worksheet.title)

    if updates:
        worksheet.update_cells(updates, value_input_option="USER_ENTERED")

    logging.info(
        "%s: %s source patients, %s updated rows, %s matched by RUT, %s matched by name, "
        "%s missing beds, %s cleared beds",
        worksheet.title,
        stats["source_patients"],
        stats["updated"],
        stats["matched_by_rut"],
        stats["matched_by_name"],
        stats["missing_bed"],
        stats["cleared_beds"],
    )
    return stats


def clear_staff_becado_for_moved_patients(
    worksheets: dict[str, gspread.Worksheet],
    schemas: dict[str, dict[str, Any]],
    moved_manual_clear_requests: list[tuple[str, int]],
    protected_manual_rows: set[tuple[str, int]],
) -> int:
    unique_requests = sorted(set(moved_manual_clear_requests))
    cleared_cells = 0

    for sheet_name, row_number in unique_requests:
        if (sheet_name, row_number) in protected_manual_rows:
            logging.info(
                "Not clearing STAFF/BECADO in %s row %s because another recognized patient now uses that row",
                sheet_name,
                row_number,
            )
            continue

        worksheet = worksheets.get(sheet_name)
        schema = schemas.get(sheet_name)
        if worksheet is None or schema is None:
            logging.warning(
                "Could not clear old STAFF/BECADO in %s row %s because the sheet is not in the target set",
                sheet_name,
                row_number,
            )
            continue

        values = schema["values"]
        columns = schema["columns"]
        updates: list[Cell] = []
        for field in STAFF_BECADO_FIELD_KEYS:
            if set_cell_if_changed(updates, values, row_number, columns.get(field), ""):
                cleared_cells += 1

        if updates:
            worksheet.update_cells(updates, value_input_option="USER_ENTERED")
            logging.info(
                "Cleared old STAFF/BECADO in %s row %s after patient moved",
                sheet_name,
                row_number,
            )

    return cleared_cells


def main() -> None:
    setup_logging()
    config = load_config()

    session = create_session()
    login_to_hospital_system(session, config.username, config.password, config.login_url)

    google_client = authorize_google_sheets(config.credentials_path)
    spreadsheet = google_client.open_by_url(config.spreadsheet_url)
    worksheets = get_target_worksheets(spreadsheet, config.target_sheet_names)

    schemas = {sheet_name: get_worksheet_schema(worksheet) for sheet_name, worksheet in worksheets.items()}
    bed_maps = {
        sheet_name: build_bed_row_map(worksheet, schemas[sheet_name])
        for sheet_name, worksheet in worksheets.items()
    }
    snapshot = build_existing_patient_snapshot(worksheets, config.target_sheet_names)

    incoming_by_sheet: dict[str, list[dict[str, str]]] = {sheet_name: [] for sheet_name in config.target_sheet_names}
    for sector in config.source_sectors:
        if sector.target_sheet not in incoming_by_sheet:
            logging.warning("Skipping sector %s; target sheet %s is not configured", sector.label, sector.target_sheet)
            continue

        df = download_sector_dataframe(
            session,
            sector.sector_id,
            config.excel_url_template,
            config.patient_list_url_template,
        )
        df = apply_sector_filter(df, sector)
        records = transform_sector_data_into_target_rows(df, sector, schemas[sector.target_sheet])
        incoming_by_sheet[sector.target_sheet].extend(records)
        logging.info(
            "Sector %s (%s) produced %s patients for %s",
            sector.label,
            sector.sector_id,
            len(records),
            sector.target_sheet,
        )

    moved_manual_clear_requests: list[tuple[str, int]] = []
    protected_manual_rows: set[tuple[str, int]] = set()

    for sheet_name in config.target_sheet_names:
        update_target_worksheet(
            worksheets[sheet_name],
            schemas[sheet_name],
            bed_maps[sheet_name],
            incoming_by_sheet[sheet_name],
            snapshot,
            config.clear_missing_beds,
            moved_manual_clear_requests,
            protected_manual_rows,
        )

    cleared_manual_cells = clear_staff_becado_for_moved_patients(
        worksheets,
        schemas,
        moved_manual_clear_requests,
        protected_manual_rows,
    )
    if cleared_manual_cells:
        logging.info("Cleared %s old STAFF/BECADO cells after patient moves", cleared_manual_cells)

    logging.info("Census update finished. No sheets outside the first 4 tabs were written.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        logging.exception("El programa fallo: %s", exc)
        if os.name == "nt":
            input("Presione Enter para cerrar esta ventana...")
        raise
