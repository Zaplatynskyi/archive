# -*- coding: utf-8 -*-
"""Витягує списки випускників із Word-документів архіву (.doc / .docx) і
створює з них Excel-описи у стандартному форматі, які далі підхоплює etl.py.

Обробляються:
  * заочна 2018–2020 — готові архівні описи (.docx), беруться повністю;
  * «Додаток до наказу про закінчення навчання» (.docx) — фільтр за
    факультетом «Механіки, енергетики та інформаційних технологій»;
  * списки дипломів (.doc/.docx, «Прийомний акт») — фільтр за переліком
    спеціальностей факультету (див. FACULTY_SPECIALTIES).

Створені файли: <Архів>/<рік>/_із_Word/<рік>_word.xlsx
               <Архів>/Заочна форма/<рік>/_із_Word/<рік>_word.xlsx

Потрібен antiword (для старих .doc). Після запуску виконайте: python etl.py
"""
from __future__ import annotations

import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict

import openpyxl

import config

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

ARCHIVE_ROOT = config.ARCHIVE_ROOT

UA_MONTHS = {
    "січня": 1, "лютого": 2, "березня": 3, "квітня": 4, "травня": 5,
    "червня": 6, "липня": 7, "серпня": 8, "вересня": 9, "жовтня": 10,
    "листопада": 11, "грудня": 12,
}

FACULTY_MARKER = "механіки, енергетики та інформаційних технологій"

# Ключові слова спеціальностей факультету (для університетських списків дипломів).
FACULTY_SPECIALTIES = [
    ("Агроінженерія", ("агроінженер", "механізаці", "процеси, машини")),
    ("Галузеве машинобудування", ("галузеве машинобуд", "машинобудуванн",
                                  "обладнання переробн", "підйомно-транспорт")),
    ("Автомобільний транспорт", ("автомобільний транспорт", "автомобілі та автомоб")),
    ("Електроенергетика, електротехніка та електромеханіка",
     ("електроенергетика", "енергетика с", "електротехн")),
    ("Інформаційні системи та технології", ("інформаційні системи",)),
    ("Комп'ютерні науки", ("комп’ютерні науки", "комп'ютерні науки")),
    ("Автоматизація та комп'ютерно-інтегровані технології",
     ("автоматизаці", "комп’ютерно-інтегров", "комп'ютерно-інтегров")),
    ("Якість, стандартизація та сертифікація", ("якість", "стандартизац", "сертифікац")),
]

HEADER = ["№ архіву", "№ за порядком", "Найменування справ (ПІБ)", "Дата початку",
          "Дата завершення", "Кількість сторінок", "Термін зберігання", "Примітка"]


# ---------------------------------------------------------------------------
# Читання документів
# ---------------------------------------------------------------------------

def docx_blocks(path):
    """Список блоків: ('P', текст) або ('ROW', [клітинки])."""
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    body = root.find(f"{W}body")
    out = []

    def walk(node):
        for el in list(node):
            tag = el.tag.replace(W, "")
            if tag == "p":
                text = "".join(t.text or "" for t in el.iter(f"{W}t"))
                if text.strip():
                    out.append(("P", text.strip()))
            elif tag == "tbl":
                for tr in el.findall(f"{W}tr"):
                    cells = []
                    for tc in tr.findall(f"{W}tc"):
                        ct = " ".join("".join(t.text or "" for t in p.iter(f"{W}t"))
                                      for p in tc.findall(f"{W}p"))
                        cells.append(ct.strip())
                    out.append(("ROW", cells))
    walk(body)
    return out


_ANTIWORD_TMP = tempfile.mkdtemp(prefix="antiword_")


def doc_lines(path):
    """Рядки старого .doc через antiword (копіюємо у ASCII-ім'я)."""
    dst = os.path.join(_ANTIWORD_TMP, "cur.doc")
    shutil.copyfile(path, dst)
    r = subprocess.run(["antiword", "-m", "UTF-8.txt", "-w", "0", dst],
                       capture_output=True)
    txt = r.stdout.decode("utf-8", "replace")
    return txt.splitlines()


# ---------------------------------------------------------------------------
# Допоміжні
# ---------------------------------------------------------------------------

def parse_ua_date(s: str) -> str | None:
    s = s.strip()
    m = re.search(r"(\d{1,2})\s+([а-яіїєґ]+)\s+(\d{4})", s.lower())
    if m and m.group(2) in UA_MONTHS:
        return f"{int(m.group(1)):02d}.{UA_MONTHS[m.group(2)]:02d}.{m.group(3)}"
    m = re.search(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})", s)
    if m:
        return f"{int(m.group(1)):02d}.{int(m.group(2)):02d}.{m.group(3)}"
    return None


def match_specialty(text: str) -> str | None:
    low = text.lower()
    for canon, keys in FACULTY_SPECIALTIES:
        if any(k in low for k in keys):
            return canon
    return None


def clean_name(*parts) -> str:
    s = " ".join(p.strip() for p in parts if p and p.strip())
    s = re.sub(r"\s+", " ", s).strip(" .,-")
    return s


NAME_RE = re.compile(r"^[А-ЯІЇЄҐ][а-яіїєґ’'ʼ\-]+(?:[ '\-][А-ЯІЇЄҐ]?[а-яіїєґ’'ʼ\-]+){1,3}$")
DIPLOMA_RE = re.compile(r"\b([МBМВM]\d{2}\s?\d{5,6}|[МВB]\d{2}\s?№?\s?\d{5,6})\b")


def looks_name(s: str) -> bool:
    s = s.strip()
    return bool(NAME_RE.match(s)) and len(s) >= 5


# ---------------------------------------------------------------------------
# Парсери конкретних форматів
# ---------------------------------------------------------------------------

def parse_opys_docx(path):
    """Готовий архівний опис (заочна 2018–2020). -> (title, [rec])."""
    blocks = docx_blocks(path)
    paras = " \n".join(t for k, t in blocks if k == "P")
    qm = re.search(r"«([^»]+)»", paras)
    dm = re.search(r"(Магістр|Бакалавр|Спеціаліст)", paras, re.I)
    spec = qm.group(1).strip() if qm else ""
    degree = dm.group(1).capitalize() if dm else ""
    title = f"{degree} — {spec}".strip(" —") or os.path.splitext(os.path.basename(path))[0]

    recs = []
    seen_header = False
    for kind, val in blocks:
        if kind != "ROW":
            continue
        cells = [c.strip() for c in val]
        joined = " ".join(cells).lower()
        if "найменування справ" in joined:
            seen_header = True
            continue
        if not seen_header:
            continue
        if "почато" in joined and "закінчено" in joined:
            continue
        # шукаємо ПІБ серед клітинок
        name = ""
        for c in cells:
            if looks_name(c):
                name = c
                break
        if not name:
            continue
        idx = cells.index(name)
        rest = cells[idx + 1:]
        dates = [c for c in rest if re.search(r"\d{2}\.\d{2}\.\d{4}", c)]
        pages = next((c for c in rest if re.fullmatch(r"\d{1,3}", c)), "")
        note = ""
        for c in rest:
            if c and c not in dates and c != pages and not re.fullmatch(r"\d+\.?", c):
                note = c
        recs.append({
            "name": name,
            "start": dates[0] if dates else "",
            "end": dates[1] if len(dates) > 1 else "",
            "pages": pages,
            "note": note,
        })
    return [(title, recs)] if recs else []


def parse_nakaz_docx(path, year):
    """«Додаток до наказу …» -> [(specialty, degree, form, [rec])]."""
    blocks = docx_blocks(path)
    groups = defaultdict(list)
    cur = None
    for kind, val in blocks:
        if kind != "ROW":
            continue
        cells = [c.strip() for c in val]
        joined = " | ".join(cells)
        low = joined.lower()
        m = re.match(r"\s*\d{3}\s+(.+?)\s+Освітня програма", joined)
        if m and FACULTY_MARKER in low:
            spec = m.group(1).strip()
            form = "заочна" if "заочна" in low else "денна"
            degree = ("Магістр" if "магістр" in low else
                      "Бакалавр" if "бакалавр" in low else
                      "Спеціаліст" if "спеціаліст" in low else None)
            cur = (spec, degree, form)
            continue
        if m:  # інший факультет
            cur = None
            continue
        if cur is None:
            continue
        if len(cells) < 5 or cells[0].lower() in ("№ з/п", ""):
            if cells[:3] == ["1", "2", "3"]:
                continue
            if not (cells and cells[0].isdigit()):
                continue
        if not cells[0].isdigit():
            continue
        # колонки: № | ЄДЕБО | Прізвище | Ім'я | По батькові | ... | Кваліфікація | Дата
        surname = cells[2] if len(cells) > 2 else ""
        first = cells[3] if len(cells) > 3 else ""
        patr = cells[4] if len(cells) > 4 else ""
        name = clean_name(surname, first, patr)
        if not name or not looks_name(name):
            continue
        end = parse_ua_date(cells[-1]) or ""
        groups[cur].append({"name": name, "start": "", "end": end,
                            "pages": "", "note": "наказ про закінчення навчання"})
    return [(s, d, f, r) for (s, d, f), r in groups.items() if r]


def parse_diploma(lines_or_blocks, year, is_docx):
    """Список дипломів («Прийомний акт»). -> [(specialty, degree, form, [rec])].
    Бере лише спеціальності факультету."""
    groups = defaultdict(list)
    degree = None
    form = None
    spec = None
    end_date = ""

    if is_docx:
        stream = []
        for kind, val in lines_or_blocks:
            if kind == "P":
                stream.append(val)
            else:
                stream.append(" | ".join(c for c in val))
    else:
        stream = lines_or_blocks

    for raw in stream:
        line = raw.strip().strip("|").strip()
        low = line.lower()
        if "освітній ступень" in low or "освітній ступінь" in low:
            degree = ("Магістр" if "магістр" in low else
                      "Бакалавр" if "бакалавр" in low else
                      "Спеціаліст" if "спеціаліст" in low else None)
            continue
        if "форма навчання" in low:
            form = "заочна" if "заочн" in low else "денна"
            continue
        if "дата закінчення" in low:
            d = parse_ua_date(line)
            if d:
                end_date = d
        if "спеціальність" in low:
            spec = match_specialty(line)
            continue
        if spec is None:
            continue
        # рядок з дипломом і ПІБ
        cells = [c.strip() for c in re.split(r"\||\t{2,}|\s{3,}", line) if c.strip()]
        dip = ""
        namebits = []
        for c in cells:
            if re.fullmatch(r"[МВBM]\s?\d{2}\s?№?\s?\d{4,6}", c) or re.match(r"^[МВBM]\d{2}\s", c):
                dip = re.sub(r"\s+", " ", c)
            elif re.search(r"[А-ЯІЇЄҐ][а-яіїєґ]", c) and "диплом" not in c.lower():
                namebits.append(c)
        name = clean_name(*namebits)
        if not name or len(name.split()) < 2 or not looks_name(name):
            continue
        key = (spec, degree, form or ("заочна" if "заочна" in str(year) else "денна"))
        groups[key].append({
            "name": name, "start": "", "end": end_date, "pages": "",
            "note": (f"диплом {dip}" if dip else "список дипломів"),
        })
    return [(s, d, f, r) for (s, d, f), r in groups.items() if r]


# ---------------------------------------------------------------------------
# Запис Excel
# ---------------------------------------------------------------------------

def sheet_name(base: str, used: set) -> str:
    name = re.sub(r"[\[\]:*?/\\]", " ", base).strip()[:31] or "Аркуш"
    n = name
    i = 2
    while n in used:
        n = f"{name[:28]}_{i}"
        i += 1
    used.add(n)
    return n


def write_workbook(out_path, sheets):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used = set()
    for title, recs in sheets:
        ws = wb.create_sheet(sheet_name(title, used))
        ws.append(["Сформовано з Word-документів архіву"])
        ws.append([])
        ws.append([])
        ws.append([])
        ws.append(HEADER)
        ws.append([title])  # рядок-роздільник зі спеціальністю (читає etl.py)
        for i, r in enumerate(recs, 1):
            ws.append(["", i, r["name"], r.get("start", ""), r.get("end", ""),
                       r.get("pages", ""), "", r.get("note", "")])
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)


# ---------------------------------------------------------------------------
# Головна логіка
# ---------------------------------------------------------------------------

def year_from_path(path):
    for part in os.path.relpath(path, ARCHIVE_ROOT).split(os.sep):
        if re.fullmatch(r"\d{4}", part):
            return int(part)
    m = re.search(r"\b(19|20)\d{2}\b", os.path.basename(path))
    return int(m.group(0)) if m else None


def dedup_key(rec, spec, form, year):
    n = re.sub(r"[^а-яіїєґ]", "", rec["name"].lower())
    return (n, year, form, spec)


def main():
    docs = []
    for ext in ("*.doc", "*.docx"):
        docs += glob.glob(os.path.join(ARCHIVE_ROOT, "**", ext), recursive=True)
    docs = sorted(d for d in docs if not os.path.basename(d).startswith("~$"))

    # (year, form) -> {sheet_title -> [rec]}
    output = defaultdict(lambda: defaultdict(list))
    seen = set()
    stats = defaultdict(int)

    for path in docs:
        rel = os.path.relpath(path, ARCHIVE_ROOT)
        year = year_from_path(path)
        if not year:
            continue
        base_form = "заочна" if "заочна форма" in rel.lower() else "денна"
        is_docx = path.lower().endswith(".docx")
        low = os.path.basename(path).lower()

        try:
            if is_docx and ("опис" in " ".join(
                    t for k, t in docx_blocks(path) if k == "P").lower()
                    or re.match(r"(бак_|.*_маг_|мех_|ен_|маш_|ат_|аі_)", low)):
                sheets = parse_opys_docx(path)
                for title, recs in sheets:
                    for r in recs:
                        k = dedup_key(r, title, base_form, year)
                        if k in seen:
                            continue
                        seen.add(k)
                        output[(year, base_form)][title].append(r)
                        stats[rel] += 1
                continue

            if is_docx and "додаток до наказу" in low:
                for spec, degree, form, recs in parse_nakaz_docx(path, year):
                    title = f"{(degree or '')} — {spec}".strip(' —')
                    for r in recs:
                        k = dedup_key(r, spec, form, year)
                        if k in seen:
                            continue
                        seen.add(k)
                        output[(year, form)][title].append(r)
                        stats[rel] += 1
                continue

            # списки дипломів
            data = docx_blocks(path) if is_docx else doc_lines(path)
            for spec, degree, form, recs in parse_diploma(data, year, is_docx):
                form = form or base_form
                title = f"{(degree or '')} — {spec}".strip(' —')
                for r in recs:
                    k = dedup_key(r, spec, form, year)
                    if k in seen:
                        continue
                    seen.add(k)
                    output[(year, form)][title].append(r)
                    stats[rel] += 1
        except Exception as e:  # noqa: BLE001
            print(f"  ! {rel}: {e}")

    total = 0
    for (year, form), sheets in sorted(output.items()):
        sub = "Заочна форма/" if form == "заочна" else ""
        out = os.path.join(ARCHIVE_ROOT, sub.replace("/", os.sep), str(year),
                           "_із_Word", f"{year}_word.xlsx")
        write_workbook(out, sorted(sheets.items()))
        n = sum(len(v) for v in sheets.values())
        total += n
        print(f"  {year} {form:7} -> {n:4} записів  ({os.path.relpath(out, ARCHIVE_ROOT)})")

    print(f"\nУсього з Word-документів: {total} записів у {len(output)} файлах.")
    print("Тепер виконайте: python etl.py")


if __name__ == "__main__":
    main()
