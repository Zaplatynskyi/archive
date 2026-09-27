"""Веб-портал архіву випускників.

Запуск для розробки:
    python app.py
Продакшн (Windows):
    waitress-serve --listen=0.0.0.0:8000 app:app
Продакшн (Linux):
    gunicorn -w 4 -b 127.0.0.1:8000 app:app
"""
from __future__ import annotations

import math
import re
import sqlite3

from flask import Flask, abort, g, render_template, request

import config

app = Flask(__name__)
app.config["SECRET_KEY"] = config.SECRET_KEY

FORMS = ["денна", "заочна"]


# ---------------------------------------------------------------------------
# Доступ до бази даних
# ---------------------------------------------------------------------------

def get_db() -> sqlite3.Connection:
    if "db" not in g:
        if not config.DB_PATH.exists():
            raise RuntimeError(
                f"База даних не знайдена: {config.DB_PATH}. "
                f"Спершу виконайте: python etl.py"
            )
        g.db = sqlite3.connect(config.DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(_exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def _meta():
    db = get_db()
    years = [r[0] for r in db.execute(
        "SELECT DISTINCT grad_year FROM graduates WHERE grad_year IS NOT NULL ORDER BY grad_year")]
    specialties = [r[0] for r in db.execute(
        "SELECT DISTINCT specialty FROM graduates WHERE specialty IS NOT NULL ORDER BY specialty")]
    degrees = [r[0] for r in db.execute(
        "SELECT DISTINCT degree FROM graduates WHERE degree IS NOT NULL ORDER BY degree")]
    total = db.execute("SELECT COUNT(*) FROM graduates").fetchone()[0]
    return {
        "years": years, "specialties": specialties, "degrees": degrees,
        "total": total, "faculty_name": config.FACULTY_NAME,
    }


@app.context_processor
def inject_meta():
    return {"faculty_name": config.FACULTY_NAME}


# ---------------------------------------------------------------------------
# Пошук
# ---------------------------------------------------------------------------

def _fts_query(raw: str) -> str:
    """Перетворює запит користувача на безпечний вираз FTS5 з префіксним пошуком."""
    tokens = re.findall(r"[\w’'ʼ\-]+", raw, flags=re.UNICODE)
    tokens = [t for t in tokens if len(t) >= 2]
    if not tokens:
        return ""
    return " AND ".join(f'"{t}"*' for t in tokens)


def search(args) -> dict:
    db = get_db()
    q = (args.get("q") or "").strip()
    form = args.get("form") or ""
    specialty = args.get("specialty") or ""
    degree = args.get("degree") or ""
    year_from = args.get("year_from", type=int)
    year_to = args.get("year_to", type=int)
    year = args.get("year", type=int)
    sort = args.get("sort") or "name"
    page = max(args.get("page", default=1, type=int), 1)

    where = []
    params: list = []
    joins = ""

    if q:
        fts = _fts_query(q)
        if fts:
            joins = "JOIN graduates_fts f ON f.rowid = g.id"
            where.append("graduates_fts MATCH ?")
            params.append(fts)
        else:
            where.append("g.full_name_raw LIKE ?")
            params.append(f"%{q}%")

    if form in FORMS:
        where.append("g.study_form = ?")
        params.append(form)
    if specialty == "__none__":
        where.append("g.specialty IS NULL")
    elif specialty:
        where.append("g.specialty = ?")
        params.append(specialty)
    if degree == "__none__":
        where.append("g.degree IS NULL")
    elif degree:
        where.append("g.degree = ?")
        params.append(degree)
    if year:
        where.append("g.grad_year = ?")
        params.append(year)
    else:
        if year_from:
            where.append("g.grad_year >= ?")
            params.append(year_from)
        if year_to:
            where.append("g.grad_year <= ?")
            params.append(year_to)

    where_sql = ("WHERE " + " AND ".join(where)) if where else ""

    order = {
        "name": "g.surname, g.first_name, g.grad_year",
        "year": "g.grad_year, g.surname, g.first_name",
        "year_desc": "g.grad_year DESC, g.surname, g.first_name",
    }.get(sort, "g.surname, g.first_name")

    total = db.execute(
        f"SELECT COUNT(*) FROM graduates g {joins} {where_sql}", params
    ).fetchone()[0]

    pages = max(math.ceil(total / config.PAGE_SIZE), 1)
    page = min(page, pages)
    offset = (page - 1) * config.PAGE_SIZE

    rows = db.execute(
        f"""SELECT g.* FROM graduates g {joins} {where_sql}
            ORDER BY {order} LIMIT ? OFFSET ?""",
        [*params, config.PAGE_SIZE, offset],
    ).fetchall()

    return {
        "rows": rows, "total": total, "page": page, "pages": pages,
        "params": {
            "q": q, "form": form, "specialty": specialty, "degree": degree,
            "year": year or "", "year_from": year_from or "",
            "year_to": year_to or "", "sort": sort,
        },
        "has_filters": bool(q or form or specialty or degree or year or year_from or year_to),
    }


# ---------------------------------------------------------------------------
# Маршрути
# ---------------------------------------------------------------------------

_DRILLDOWN_SKIP_KEYS = ("q", "year_from", "year_to", "degree", "page", "sort", "all")


def _is_browsing(args) -> bool:
    """Просте клацання по року (лише ?year=NNNN) веде в покроковий вибір
    форми/напряму. Будь-який інший параметр (пошук, сортування, сторінка
    пейджера, явний скид фільтрів...) означає, що користувач уже працює
    з формою пошуку — тоді одразу показуємо плаский список результатів."""
    if any(args.get(k) for k in _DRILLDOWN_SKIP_KEYS):
        return False
    return bool(args.get("year", type=int))


@app.route("/")
def index():
    meta = _meta()
    if request.args:
        if _is_browsing(request.args):
            year = request.args.get("year", type=int)
            form = request.args.get("form") or ""
            if not form:
                return render_template(
                    "choose_form.html", meta=meta, year=year, forms=_form_stats(year)
                )
            if form in FORMS and not request.args.get("specialty"):
                return render_template(
                    "choose_specialty.html", meta=meta, year=year, form=form,
                    specialties=_specialty_stats(year, form),
                )
            specialty = request.args.get("specialty") or ""
            if form in FORMS and specialty:
                degrees = _degree_stats(year, form, specialty)
                if len(degrees) > 1:
                    return render_template(
                        "choose_degree.html", meta=meta, year=year, form=form,
                        specialty=specialty, degrees=degrees,
                    )
        result = search(request.args)
        return render_template("results.html", meta=meta, **result)
    stats = _year_stats()
    return render_template("index.html", meta=meta, stats=stats)


def _year_stats():
    db = get_db()
    return db.execute(
        """SELECT grad_year,
                  SUM(study_form='денна')  AS denna,
                  SUM(study_form='заочна') AS zaochna,
                  COUNT(*) AS total
           FROM graduates WHERE grad_year IS NOT NULL
           GROUP BY grad_year ORDER BY grad_year"""
    ).fetchall()


def _form_stats(year: int):
    db = get_db()
    return db.execute(
        """SELECT study_form, COUNT(*) AS total
           FROM graduates WHERE grad_year = ?
           GROUP BY study_form ORDER BY study_form""",
        (year,),
    ).fetchall()


def _specialty_stats(year: int, form: str):
    db = get_db()
    return db.execute(
        """SELECT specialty, COUNT(*) AS total
           FROM graduates WHERE grad_year = ? AND study_form = ?
           GROUP BY specialty ORDER BY specialty IS NULL, specialty""",
        (year, form),
    ).fetchall()


_DEGREE_ORDER = (
    "Молодший спеціаліст", "Бакалавр", "Спеціаліст", "Інженер", "Магістр",
)


def _degree_stats(year: int, form: str, specialty: str):
    db = get_db()
    where = "grad_year = ? AND study_form = ?"
    params: list = [year, form]
    if specialty == "__none__":
        where += " AND specialty IS NULL"
    else:
        where += " AND specialty = ?"
        params.append(specialty)
    rows = db.execute(
        f"SELECT degree, COUNT(*) AS total FROM graduates WHERE {where} GROUP BY degree",
        params,
    ).fetchall()

    def sort_key(r):
        d = r["degree"]
        if d is None:
            return (2, "")
        try:
            return (0, _DEGREE_ORDER.index(d))
        except ValueError:
            return (1, d)

    return sorted(rows, key=sort_key)


@app.route("/graduate/<int:gid>")
def graduate(gid: int):
    db = get_db()
    row = db.execute("SELECT * FROM graduates WHERE id = ?", (gid,)).fetchone()
    if row is None:
        abort(404)
    same_year = db.execute(
        """SELECT id, full_name FROM graduates
           WHERE grad_year = ? AND study_form = ? AND id != ?
           ORDER BY surname, first_name LIMIT 400""",
        (row["grad_year"], row["study_form"], gid),
    ).fetchall()
    return render_template("graduate.html", g=row, same_year=same_year, meta=_meta())


@app.route("/about")
def about():
    db = get_db()
    by_year = _year_stats()
    covered = [r["grad_year"] for r in by_year]
    return render_template(
        "about.html", meta=_meta(), by_year=by_year,
        year_min=covered[0] if covered else None,
        year_max=covered[-1] if covered else None,
    )


@app.errorhandler(404)
def not_found(_e):
    return render_template("404.html", meta=_meta()), 404


if __name__ == "__main__":
    app.run(debug=True, port=5000)
