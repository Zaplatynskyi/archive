"""Імпорт архіву випускників з Excel-файлів у базу даних SQLite.

Запуск:
    python etl.py                     # використати теку з config.ARCHIVE_ROOT
    python etl.py "D:\\шлях\\до\\архіву"  # вказати теку вручну

Скрипт повністю перебудовує базу даних (data/alumni.db) з нуля, тож його
безпечно запускати повторно після оновлення Excel-файлів.
"""
from __future__ import annotations

import re
import sqlite3
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path

import openpyxl

import config

# ---------------------------------------------------------------------------
# Розпізнавання колонок
# ---------------------------------------------------------------------------

YEAR_RE = re.compile(r"^\d{4}$")
DIGIT_RE = re.compile(r"\d")


def _norm(text) -> str:
    if text is None:
        return ""
    s = str(text).replace("\n", " ").replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def classify_columns(header_row) -> dict:
    """За рядком-заголовком повертає відображення {роль: індекс колонки}."""
    cols: dict[str, int] = {}
    cells = [_norm(c) for c in header_row]
    for i, h in enumerate(cells):
        if not h:
            continue
        if ("наймену" in h or "назва справ" in h) and "name" not in cols:
            cols["name"] = i
        elif ("архів" in h or "архир" in h) and "archive_no" not in cols:
            cols["archive_no"] = i
        elif ("п/п" in h or "п.п." in h or "за порядк" in h or "за порядок" in h) and "seq" not in cols:
            cols["seq"] = i
        elif ("початку" in h or "початок" in h or "почато" in h) and "start" not in cols:
            cols["start"] = i
        elif ("завершенн" in h or "закінч" in h) and "end" not in cols:
            cols["end"] = i
        elif ("кільк" in h or "к-сть" in h) and "pages" not in cols:
            cols["pages"] = i
        elif "зберіганн" in h and "retention" not in cols:
            cols["retention"] = i
        elif "приміт" in h and "note" not in cols:
            cols["note"] = i
        elif h == "дата" and "start" not in cols:
            # об'єднаний заголовок «Дата» над двома колонками (почато/закінчено)
            cols["start"] = i
            cols["end"] = i + 1
    return cols


def find_header(rows) -> tuple[int, dict] | tuple[None, None]:
    for i, r in enumerate(rows[:40]):
        joined = " ".join(_norm(c) for c in r)
        if "наймену" in joined or "назва справ" in joined:
            cols = classify_columns(r)
            if "name" in cols:
                return i, cols
    return None, None


# ---------------------------------------------------------------------------
# Розбір значень
# ---------------------------------------------------------------------------

_DATE_FORMATS = [
    "%d.%m.%Y", "%d.%m.%y", "%d/%m/%Y", "%d/%m/%y",
    "%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%d-%m-%Y", "%d-%m-%y",
]


def parse_date(value) -> tuple[str, int | None]:
    """Повертає (сирий_текст, рік_або_None)."""
    if value is None:
        return "", None
    if isinstance(value, (datetime, date)):
        return value.strftime("%d.%m.%Y"), value.year
    s = str(value).strip()
    if not s:
        return "", None
    for fmt in _DATE_FORMATS:
        try:
            d = datetime.strptime(s, fmt)
            y = d.year
            if 1930 <= y <= 2035:
                return s, y
        except ValueError:
            pass
    # шаблони на кшталт «16/3-61», «1/7-67»
    m = re.search(r"(\d{1,2})[./\-](\d{1,2})[./\-](\d{2,4})", s)
    if m:
        y = int(m.group(3))
        if y < 100:
            y += 1900 if y > 30 else 2000
        if 1930 <= y <= 2035:
            return s, y
    m = re.search(r"(19|20)\d{2}", s)
    if m and 1930 <= int(m.group(0)) <= 2035:
        return s, int(m.group(0))
    return s, None


_DITTO_RE = re.compile(r"^[-\"'«»—–\s.,_]+$")


def is_ditto(value) -> bool:
    if value is None:
        return False
    s = str(value).strip().lower()
    if not s:
        return False
    return bool(_DITTO_RE.match(s)) or s in {"те саме", "той самий", "-“-", "–“–"}


def clean_archive_no(value) -> str | None:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    m = re.fullmatch(r"(\d+)\.0+", s)
    return m.group(1) if m else s


def parse_int(value) -> int | None:
    if value is None:
        return None
    try:
        return int(float(str(value).replace(",", ".").strip()))
    except (ValueError, TypeError):
        return None


# ---------------------------------------------------------------------------
# Імена
# ---------------------------------------------------------------------------

CYRILLIC_RE = re.compile(r"[А-Яа-яІіЇїЄєҐґ'’ʼ\-]")
JUNK_STARTS = (
    "опис", "запис", "львівський", "факультет", "перелік", "продовження",
    "разом", "всього", "усього", "аркуш", "таблиц", "додаток", "спеціальн",
    "магістр", "бакалавр", "спеціаліст", "наймену", "назва справ", "прізвище",
    "п.і.б", "піб", "№",
)


def looks_like_name(s: str) -> bool:
    s = s.strip()
    if len(s) < 4:
        return False
    low = s.lower()
    if low.startswith(JUNK_STARTS):
        return False
    if "наймену" in low or "найменування справ" in low:
        return False
    if not CYRILLIC_RE.search(s):
        return False
    tokens = [t for t in re.split(r"\s+", s) if t]
    return len(tokens) >= 2


def split_name(s: str) -> tuple[str, str, str, str]:
    """Повертає (повне_відображуване, прізвище, ім'я, по-батькові)."""
    s = re.sub(r"\s+", " ", s.replace("\xa0", " ")).strip(" .,;")

    def nice(tok: str) -> str:
        if not tok:
            return ""
        parts = re.split(r"([-()\s])", tok)
        out = []
        for p in parts:
            if p and p.isalpha():
                out.append(p[:1].upper() + p[1:].lower())
            else:
                out.append(p)
        return "".join(out)

    display = " ".join(nice(t) for t in s.split(" "))

    # прізвище/ім'я/по батькові — без дужкових приміток, для індексації
    core = re.sub(r"\s*\([^)]*\)\s*", " ", s).strip()
    tokens = [t for t in core.split(" ") if t]
    surname = nice(tokens[0]) if tokens else ""
    first = nice(tokens[1]) if len(tokens) > 1 else ""
    patr = " ".join(nice(t) for t in tokens[2:]) if len(tokens) > 2 else ""
    return display or s, surname, first, patr


# ---------------------------------------------------------------------------
# Спеціальність / освітній рівень
# ---------------------------------------------------------------------------

def classify_specialty(*texts: str) -> tuple[str | None, str | None]:
    blob = " ".join(_norm(t) for t in texts if t)
    degree = None
    if "магістр" in blob:
        degree = "Магістр"
    elif "бакалавр" in blob:
        degree = "Бакалавр"
    elif "спеціаліст" in blob:
        degree = "Спеціаліст"
    elif "інженер" in blob:
        degree = "Інженер"

    specialty = None
    if any(k in blob for k in ("якіст", "стандарт", "сертиф")):
        specialty = "Якість, стандартизація та сертифікація"
    elif any(k in blob for k in (
        "інформаційн", "комп’ютер", "комп'ютер", "автоматизац", "програмн",
        "робототехн",
    )):
        specialty = "Інформаційні технології"
    elif any(k in blob for k in ("енергет", "електро")):
        specialty = "Енергетика та електротехніка"
    elif any(k in blob for k in ("автомоб", "автотранспорт", "автомобільний транспорт")):
        specialty = "Автомобільний транспорт"
    elif any(k in blob for k in ("агроінженер", "агроінж")):
        specialty = "Агроінженерія"
    elif any(k in blob for k in (
        "механіз", "механік", "машин", "транспорт", "процеси",
        "обладн", "підйомно", "харчов", "переробн",
    )):
        specialty = "Механізація та машинобудування"
    elif any(k in blob for k in ("агро", "апк")):
        specialty = "Агроінженерія"
    return specialty, degree


# ---------------------------------------------------------------------------
# Обробка файлів
# ---------------------------------------------------------------------------

def year_and_form(rel_parts) -> tuple[int | None, str]:
    form = "заочна" if any("заочна форма" in _norm(p) for p in rel_parts) else "денна"
    year = None
    for p in rel_parts:
        if YEAR_RE.match(p.strip()):
            year = int(p.strip())
            break
    return year, form


def process_workbook(path: Path, archive_root: Path, records: list, stats: dict):
    rel = path.relative_to(archive_root)
    grad_year, form = year_and_form(rel.parts)
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:  # noqa: BLE001
        stats["bad_files"].append(f"{rel}: {e}")
        return

    for ws in wb.worksheets:
        rows = []
        for i, r in enumerate(ws.iter_rows(values_only=True)):
            rows.append(r)
            if i > 5000:
                break
        if not any(any(c is not None for c in r) for r in rows):
            continue
        hidx, cols = find_header(rows)
        if hidx is None:
            stats["sheets_no_header"] += 1
            continue

        sheet_spec, sheet_degree = classify_specialty(ws.title, path.stem)
        current_section = None
        name_i = cols["name"]
        prev = {"start": None, "end": None, "retention": None}

        start_data = hidx + 1
        # пропустити підзаголовок «почато / закінчено»
        if start_data < len(rows):
            sub = " ".join(_norm(c) for c in rows[start_data])
            if sub and ("почато" in sub or "закінч" in sub or "початок" in sub) and "наймену" not in sub:
                start_data += 1

        for r in rows[start_data:]:
            if name_i >= len(r):
                continue
            raw = r[name_i]
            non_empty = [c for c in r if c is not None and str(c).strip() != ""]
            if raw is None or str(raw).strip() == "":
                # можливий рядок-роздільник зі спеціальністю
                if len(non_empty) == 1 and not DIGIT_RE.search(str(non_empty[0])):
                    current_section = str(non_empty[0]).strip()
                continue
            s = str(raw).strip()

            if len(non_empty) == 1 and not DIGIT_RE.search(s):
                current_section = s
                continue
            if not looks_like_name(s):
                stats["rows_skipped"] += 1
                continue

            def cell(role):
                idx = cols.get(role)
                if idx is None or idx >= len(r):
                    return None
                return r[idx]

            display, surname, first, patr = split_name(s)

            start_val = cell("start")
            if is_ditto(start_val):
                start_val = prev["start"]
            elif start_val is not None and str(start_val).strip():
                prev["start"] = start_val
            end_val = cell("end")
            if is_ditto(end_val):
                end_val = prev["end"]
            elif end_val is not None and str(end_val).strip():
                prev["end"] = end_val
            retention_val = cell("retention")
            if is_ditto(retention_val):
                retention_val = prev["retention"]
            elif retention_val is not None and str(retention_val).strip():
                prev["retention"] = retention_val

            start_raw, start_year = parse_date(start_val)
            end_raw, end_year = parse_date(end_val)

            spec, degree = classify_specialty(
                current_section or "", ws.title, path.stem
            )
            spec = spec or sheet_spec
            degree = degree or sheet_degree

            records.append(
                {
                    "full_name": display,
                    "full_name_raw": re.sub(r"\s+", " ", s),
                    "surname": surname,
                    "first_name": first,
                    "patronymic": patr,
                    "grad_year": grad_year or end_year,
                    "study_form": form,
                    "specialty": spec,
                    "degree": degree,
                    "section": current_section,
                    "archive_no": clean_archive_no(cell("archive_no")),
                    "seq_no": parse_int(cell("seq")),
                    "start_date": start_raw or None,
                    "end_date": end_raw or None,
                    "start_year": start_year,
                    "end_year": end_year,
                    "pages": parse_int(cell("pages")),
                    "retention": (str(retention_val).strip()
                                  if retention_val is not None and str(retention_val).strip() else None),
                    "note": (str(cell("note")).strip()
                             if cell("note") is not None else None),
                    "source_file": str(rel).replace("\\", "/"),
                    "source_sheet": ws.title,
                }
            )
        stats["sheets_ok"] += 1
    wb.close()


# ---------------------------------------------------------------------------
# База даних
# ---------------------------------------------------------------------------

SCHEMA = """
DROP TABLE IF EXISTS graduates;
DROP TABLE IF EXISTS graduates_fts;

CREATE TABLE graduates (
    id           INTEGER PRIMARY KEY,
    full_name    TEXT NOT NULL,
    full_name_raw TEXT NOT NULL,
    surname      TEXT,
    first_name   TEXT,
    patronymic   TEXT,
    grad_year    INTEGER,
    study_form   TEXT,
    specialty    TEXT,
    degree       TEXT,
    section      TEXT,
    archive_no   TEXT,
    seq_no       INTEGER,
    start_date   TEXT,
    end_date     TEXT,
    start_year   INTEGER,
    end_year     INTEGER,
    pages        INTEGER,
    retention    TEXT,
    note         TEXT,
    source_file  TEXT,
    source_sheet TEXT
);

CREATE INDEX idx_grad_year ON graduates(grad_year);
CREATE INDEX idx_form ON graduates(study_form);
CREATE INDEX idx_specialty ON graduates(specialty);
CREATE INDEX idx_surname ON graduates(surname);

CREATE VIRTUAL TABLE graduates_fts USING fts5(
    full_name, full_name_raw, specialty, note, section,
    content='graduates', content_rowid='id',
    tokenize="unicode61 remove_diacritics 0"
);
"""


def normalize_key(rec) -> tuple:
    name = re.sub(r"[^а-яіїєґ]", "", rec["full_name_raw"].lower())
    return (name, rec["grad_year"], rec["study_form"])


def build_database(archive_root: Path, db_path: Path) -> dict:
    archive_root = Path(archive_root)
    if not archive_root.exists():
        raise SystemExit(f"Тека архіву не знайдена: {archive_root}")

    files = sorted(archive_root.rglob("*.xlsx"))
    files = [f for f in files if not f.name.startswith("~$")]
    stats = {
        "files": len(files),
        "sheets_ok": 0,
        "sheets_no_header": 0,
        "rows_skipped": 0,
        "bad_files": [],
    }
    records: list[dict] = []
    for f in files:
        process_workbook(f, archive_root, records, stats)

    # дедуплікація: один випускник = (нормалізоване ПІБ, рік, форма)
    seen: dict[tuple, dict] = {}
    for rec in records:
        key = normalize_key(rec)
        if None in key[1:]:
            key = key + (rec["source_file"], rec["archive_no"], id(rec))
        if key in seen:
            prev = seen[key]
            merged = prev.setdefault("_sources", {prev["source_file"]})
            merged.add(rec["source_file"])
            for field in ("archive_no", "start_date", "end_date", "pages",
                          "retention", "note", "specialty", "degree", "section"):
                if not prev.get(field) and rec.get(field):
                    prev[field] = rec[field]
            continue
        seen[key] = rec

    unique = list(seen.values())
    stats["records_raw"] = len(records)
    stats["records_unique"] = len(unique)

    db_path.parent.mkdir(parents=True, exist_ok=True)
    if db_path.exists():
        db_path.unlink()
    con = sqlite3.connect(db_path)
    con.executescript(SCHEMA)
    cur = con.cursor()
    for rec in unique:
        srcs = rec.get("_sources")
        if srcs:
            rec = dict(rec)
            rec["source_file"] = "; ".join(sorted(srcs))
        cur.execute(
            """INSERT INTO graduates
               (full_name, full_name_raw, surname, first_name, patronymic,
                grad_year, study_form, specialty, degree, section, archive_no,
                seq_no, start_date, end_date, start_year, end_year, pages,
                retention, note, source_file, source_sheet)
               VALUES (:full_name, :full_name_raw, :surname, :first_name,
                :patronymic, :grad_year, :study_form, :specialty, :degree,
                :section, :archive_no, :seq_no, :start_date, :end_date,
                :start_year, :end_year, :pages, :retention, :note,
                :source_file, :source_sheet)""",
            {k: rec.get(k) for k in (
                "full_name", "full_name_raw", "surname", "first_name",
                "patronymic", "grad_year", "study_form", "specialty", "degree",
                "section", "archive_no", "seq_no", "start_date", "end_date",
                "start_year", "end_year", "pages", "retention", "note",
                "source_file", "source_sheet")},
        )
    con.execute(
        "INSERT INTO graduates_fts(rowid, full_name, full_name_raw, specialty, note, section) "
        "SELECT id, full_name, full_name_raw, specialty, note, section FROM graduates"
    )
    con.commit()

    stats["years"] = [r[0] for r in con.execute(
        "SELECT DISTINCT grad_year FROM graduates WHERE grad_year IS NOT NULL ORDER BY grad_year")]
    con.close()
    return stats


def main():
    root = Path(sys.argv[1]) if len(sys.argv) > 1 else config.ARCHIVE_ROOT
    print(f"Архів: {root}")
    print(f"База даних: {config.DB_PATH}")
    stats = build_database(root, config.DB_PATH)
    print("\n=== Підсумок імпорту ===")
    print(f"  Excel-файлів опрацьовано:      {stats['files']}")
    print(f"  Аркушів з даними:              {stats['sheets_ok']}")
    print(f"  Аркушів без заголовка:         {stats['sheets_no_header']}")
    print(f"  Рядків відкинуто (не ПІБ):     {stats['rows_skipped']}")
    print(f"  Записів усього:                {stats['records_raw']}")
    print(f"  Унікальних випускників:        {stats['records_unique']}")
    if stats["years"]:
        print(f"  Роки випуску:                  {stats['years'][0]}–{stats['years'][-1]}")
    if stats["bad_files"]:
        print("  Проблемні файли:")
        for b in stats["bad_files"]:
            print(f"    - {b}")


if __name__ == "__main__":
    main()
