"""Конфігурація порталу. Значення можна перевизначити через змінні середовища
або файл .env (див. .env.example)."""
import os
from pathlib import Path

try:
    from dotenv import load_dotenv

    load_dotenv()
except Exception:  # python-dotenv не обов'язковий
    pass

BASE_DIR = Path(__file__).resolve().parent

# Тека з архівом (68 річних тек + «Заочна форма»). За замовчуванням шукаємо
# поряд із проєктом.
ARCHIVE_ROOT = Path(
    os.environ.get("ARCHIVE_ROOT", BASE_DIR / "Архів випускників")
).resolve()

# Файл бази даних SQLite.
DB_PATH = Path(os.environ.get("DB_PATH", BASE_DIR / "data" / "alumni.db")).resolve()

# Назва підрозділу для шапки сайту.
FACULTY_NAME = os.environ.get(
    "FACULTY_NAME",
    "Факультет механіки, енергетики та інформаційних технологій",
)

SECRET_KEY = os.environ.get("SECRET_KEY", "dev-key-change-me")

# Кількість записів на сторінці результатів.
PAGE_SIZE = int(os.environ.get("PAGE_SIZE", "40"))
