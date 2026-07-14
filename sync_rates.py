"""
Безопасная синхронизация тарифов в лист «Ставки» таблицы руководства.

Трогает ТОЛЬКО лист «Ставки»:
  - перезаписывает тарифные ставки (берёт TARIFF_RATES из setup_mgmt_sheet.py);
  - сохраняет список сотрудников: читает существующих из листа, добавляет
    новых из EMPLOYEES, отсутствующих в листе;
  - применяет форматирование (как setup_mgmt_sheet.setup_rates_sheet).

НЕ ТРОГАЕТ: «История_ЗП», «Корректировки», «Итого_к_выплате».

Запуск:
  python sync_rates.py

Учётные данные: credentials.json рядом (CREDENTIALS_FILE из config.py)
или переменная окружения GOOGLE_CREDENTIALS_B64.
"""

import os
import sys
import base64
import json
import time

import gspread
from google.oauth2.service_account import Credentials

try:
    from config import MGMT_SHEETS_URL, CREDENTIALS_FILE
except ImportError:
    MGMT_SHEETS_URL = os.environ.get("MGMT_SHEETS_URL", "")
    CREDENTIALS_FILE = "credentials.json"

import setup_mgmt_sheet as s
from setup_mgmt_sheet import _with_retry

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _creds_path():
    base = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base, CREDENTIALS_FILE)
    if os.path.exists(path):
        return path
    # Fallback: GOOGLE_CREDENTIALS_B64
    b64 = os.environ.get("GOOGLE_CREDENTIALS_B64", "")
    if b64:
        data = base64.b64decode(b64).decode("utf-8")
        json.loads(data)  # проверка, что это валидный JSON
        tmp = os.path.join(base, "_credentials_cloud.json")
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(data)
        return tmp
    raise FileNotFoundError(
        f"Учётные данные не найдены: положи {CREDENTIALS_FILE} рядом со скриптом "
        "или задай GOOGLE_CREDENTIALS_B64."
    )


def get_client():
    creds = Credentials.from_service_account_file(_creds_path(), scopes=SCOPES)
    try:
        print(f"  Сервисный аккаунт: {creds.service_account_email}")
    except Exception:
        pass
    return gspread.authorize(creds)


def read_existing_employees(ws):
    """Читает раздел «СОТРУДНИКИ» из листа Ставки → список [Имя, Роль, Коэф, Активен]."""
    rows = _with_retry(ws.get_all_values)
    emp = []
    in_emp = False
    for row in rows:
        first = row[0].strip() if row and row[0] else ""
        if first == "Имя":
            in_emp = True
            continue
        if not in_emp:
            continue
        if not row or not any(c.strip() for c in row):
            break
        if first.startswith("*"):
            break
        if not first:
            continue
        r = list(row) + [""] * 4
        emp.append([r[0].strip(), r[1].strip(), r[2].strip() or "1.0", r[3].strip() or "ДА"])
    return emp


def merge_employees(existing, from_script):
    """Объединяет: существующие остаются как есть, из скрипта добавляются только
    те, кого ещё нет в листе (по нормализованному имени)."""
    def key(name):
        return name.strip().lower()
    merged = list(existing)
    seen = {key(e[0]) for e in existing}
    for e in from_script:
        if key(e[0]) not in seen:
            merged.append(list(e))
            seen.add(key(e[0]))
    return merged


def main():
    print("=" * 60)
    print("  SignaturePro — Синхронизация тарифов (только лист «Ставки»)")
    print("=" * 60)

    if not MGMT_SHEETS_URL:
        print("  Ошибка: MGMT_SHEETS_URL не задан в config.py")
        sys.exit(1)

    print("\nПодключение к Google Sheets API...")
    gc = get_client()

    print("Открываю таблицу руководства...")
    sh = _with_retry(lambda: gc.open_by_url(MGMT_SHEETS_URL))
    print(f"  Открыта: {sh.title}")

    existing_titles = {ws.title for ws in sh.worksheets()}

    if "Ставки" in existing_titles:
        ws = sh.worksheet("Ставки")
        print("  Лист «Ставки» найден — сохраняю существующих сотрудников...")
        existing_emp = read_existing_employees(ws)
        print(f"  Найдено сотрудников в листе: {len(existing_emp)}")
    else:
        print("  Лист «Ставки» не найден — создаю...")
        ws = _with_retry(lambda: sh.add_worksheet(title="Ставки", rows=50, cols=10))
        existing_emp = []

    merged_emp = merge_employees(existing_emp, s.EMPLOYEES)
    print(f"  Итого сотрудников после слияния: {len(merged_emp)}")
    if len(merged_emp) > len(existing_emp):
        print(f"  Добавлено новых из скрипта: {len(merged_emp) - len(existing_emp)}")

    # Подменяем глобальный список сотрудников и перерисовываем лист «Ставки».
    # setup_rates_sheet очищает ТОЛЬКО лист «Ставки» и записывает тарифы + сотрудников.
    s.EMPLOYEES = merged_emp
    s.setup_rates_sheet(sh, ws)

    # Печатаем сводку по тарифам
    print("\n  Тарифы в листе «Ставки» после синхронизации:")
    for t in s.TARIFF_RATES:
        print(f"    {t[0]:<14} вар={t[1]:>4} об={t[2]:>4} бп={t[3]:>4} "
              f"тб={t[4]:>4} пр={t[5]:>4} мен={t[6]:>4}")

    print("\n" + "=" * 60)
    print("  [DONE] Лист «Ставки» обновлён.")
    print("  История_ЗП и Корректировки НЕ затронуты.")
    print("  Дашборд подхватит новые ставки в течение 10 минут (кеш),")
    print("  или сразу — кнопкой «Обновить».")
    print("=" * 60)


if __name__ == "__main__":
    main()