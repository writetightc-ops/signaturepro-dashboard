"""
Адаптер для чтения/записи данных из Google Таблицы через Google Sheets API.
Использует gspread + google-auth (сервисный аккаунт).
"""

import time
import os
from datetime import datetime, date

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Листы, которые не являются таблицами заказов
SKIP_SHEETS = {
    "Справочник", "Шаблон", "import", "export", "Дисциплина",
    "ЗП отчет месяц", "Бонусы", "Рейты", "Сотрудники",
    "Статистика", "Импорт", "Sheet1", "Лист1", "Sheet",
}

_client = None   # переиспользуемый gspread-клиент
_cache = {}      # {url: {"ts": float, "orders": list, "sheets": list}}


# ─── Авторизация ─────────────────────────────────────────────────────────────

def _get_client(credentials_file):
    global _client
    if _client is not None:
        return _client
    if not GSPREAD_AVAILABLE:
        raise ImportError("Установи библиотеки: pip install gspread google-auth")
    if not os.path.exists(credentials_file):
        raise FileNotFoundError(
            f"Файл учётных данных не найден: {credentials_file}\n"
            "Скачай его из Google Cloud Console (Service Account → Keys → JSON)"
        )
    creds = Credentials.from_service_account_file(credentials_file, scopes=SCOPES)
    _client = gspread.authorize(creds)
    return _client


def _with_retry(fn, max_retries=6):
    """Повторяет запрос при ошибке лимита API (429 Too Many Requests)."""
    last_err = None
    for attempt in range(max_retries):
        try:
            return fn()
        except Exception as e:
            msg = str(e)
            is_rate_limit = "429" in msg or "RESOURCE_EXHAUSTED" in msg or "quota" in msg.lower()
            if is_rate_limit:
                # Начинаем с 15с (write-лимит = 60с), экспоненциально растём
                wait = 15 * (2 ** attempt)  # 15, 30, 60, 120, 240, 480 сек
                wait = min(wait, 120)  # не ждём больше 2 минут за раз
                print(f"Google Sheets API: лимит запросов, жду {wait}с (попытка {attempt+1}/{max_retries})...")
                time.sleep(wait)
                last_err = e
            else:
                raise
    raise Exception(f"Превышен лимит запросов Google Sheets API после {max_retries} попыток: {last_err}")


# ─── Парсинг значений из Google Sheets (всё приходит строками) ───────────────

def _parse_date(val):
    """Преобразует строку или datetime/date объект в date."""
    if val is None or val == "":
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if not isinstance(val, str):
        return None
    s = val.strip()
    if not s:
        return None
    for fmt in ["%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d.%m.%y",
                "%m/%d/%Y", "%d-%m-%Y", "%d-%m-%y"]:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _parse_bool(val):
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return bool(val)
    if isinstance(val, str):
        return val.strip().upper() in ("TRUE", "ДА", "YES", "1", "ИСТИНА", "+")
    return False


def _row_to_order(row, sheet_name):
    """Преобразует строку листа (список значений) в словарь заказа."""
    r = list(row) + [""] * 25  # на случай если строка короткая
    # Пустая строка — вернём None
    if not r[0] and not r[2]:
        return None
    # Строка-подсказка из шаблона
    if isinstance(r[0], str) and r[0].startswith("←"):
        return None
    return {
        "поступление": _parse_date(r[0]),
        "статус":      _parse_bool(r[1]),
        "менеджер":    str(r[2]).strip() if r[2] else "",
        "каллиграф":   str(r[3]).strip() if r[3] else "",
        "клиент":      str(r[4]).strip() if r[4] else "",
        "тариф":       str(r[5]).strip().upper() if r[5] else "",
        "варианты":    _parse_date(r[6]),
        "ос1":         _parse_date(r[7]),
        "правка1":     _parse_date(r[8]),
        "ос2":         _parse_date(r[9]),
        "правка2":     _parse_date(r[10]),
        "ос3":         _parse_date(r[11]),
        "правка3":     _parse_date(r[12]),
        "ос4":         _parse_date(r[13]),
        "правка4":     _parse_date(r[14]),
        "обучение":    _parse_date(r[15]),
        "ссылка":      str(r[16]).strip() if r[16] else "",
        "завершение":  _parse_date(r[17]),
        "заметки":     str(r[18]).strip() if len(r) > 18 and r[18] else "",
        "_sheet":      sheet_name,
    }


# ─── Основные функции ─────────────────────────────────────────────────────────

def get_order_sheet_names(spreadsheet_url, credentials_file):
    """Возвращает список листов заказов (исключая служебные)."""
    client = _get_client(credentials_file)
    sh = _with_retry(lambda: client.open_by_url(spreadsheet_url))
    return [ws.title for ws in sh.worksheets() if ws.title not in SKIP_SHEETS]


def read_all_orders(spreadsheet_url, credentials_file, cache_ttl=45):
    """
    Читает все заказы из всех листов заказов.
    Возвращает (orders: list[dict], sheets_used: list[str]).
    Кеширует результат на cache_ttl секунд.
    """
    key = spreadsheet_url
    now = time.time()
    if key in _cache and (now - _cache[key]["ts"]) < cache_ttl:
        return _cache[key]["orders"], _cache[key]["sheets"]

    client = _get_client(credentials_file)
    sh = _with_retry(lambda: client.open_by_url(spreadsheet_url))
    worksheets = [ws for ws in sh.worksheets() if ws.title not in SKIP_SHEETS]

    all_orders = []
    sheets_used = []

    for ws in worksheets:
        rows = _with_retry(ws.get_all_values)
        if len(rows) < 2:
            continue
        sheets_used.append(ws.title)
        for row in rows[1:]:  # пропускаем заголовок
            order = _row_to_order(row, ws.title)
            if order:
                all_orders.append(order)

    _cache[key] = {"ts": now, "orders": all_orders, "sheets": sheets_used}
    return all_orders, sheets_used


def invalidate_cache(spreadsheet_url):
    """Сбрасывает кеш после изменения данных (перенос, запись)."""
    if spreadsheet_url in _cache:
        del _cache[spreadsheet_url]


def transfer_orders(spreadsheet_url, credentials_file, from_sheet_name, to_sheet_name):
    """
    Переносит незавершённые заказы (без даты завершения) из одного листа в другой.
    Возвращает количество перенесённых строк.
    """
    client = _get_client(credentials_file)
    sh = _with_retry(lambda: client.open_by_url(spreadsheet_url))

    # Исходный лист
    try:
        ws_from = sh.worksheet(from_sheet_name)
    except Exception:
        raise ValueError(f"Лист «{from_sheet_name}» не найден в таблице")

    # Целевой лист — создаём, если нет
    try:
        ws_to = sh.worksheet(to_sheet_name)
    except Exception:
        ws_to = _with_retry(lambda: sh.add_worksheet(title=to_sheet_name, rows=500, cols=19))
        # Копируем заголовки
        headers = _with_retry(lambda: ws_from.row_values(1))
        _with_retry(lambda: ws_to.append_row(headers, value_input_option="USER_ENTERED"))

    all_rows = _with_retry(ws_from.get_all_values)
    if len(all_rows) < 2:
        return 0

    headers   = all_rows[0]
    data_rows = all_rows[1:]

    DONE_COL = 17  # 0-based индекс «Дата завершения»

    rows_to_move = []
    rows_to_keep = [headers]

    for row in data_rows:
        r = row + [""] * 5
        if not r[0] and not r[2]:
            continue  # пустая строка — пропуск
        done = r[DONE_COL].strip() if DONE_COL < len(r) else ""
        if not done:
            rows_to_move.append(row)
        else:
            rows_to_keep.append(row)

    if not rows_to_move:
        return 0

    # Добавляем в целевой лист одним батчевым запросом
    padded_rows = [(row + [""] * 19)[:19] for row in rows_to_move]
    _with_retry(lambda: ws_to.append_rows(padded_rows, value_input_option="USER_ENTERED"))

    # Небольшая пауза между write-операциями, чтобы не переполнить квоту
    time.sleep(2)

    # Перезаписываем исходный лист (только незавершённые уходят)
    _with_retry(lambda: ws_from.clear())
    time.sleep(1)
    if rows_to_keep:
        _with_retry(lambda: ws_from.update("A1", rows_to_keep, value_input_option="USER_ENTERED"))

    invalidate_cache(spreadsheet_url)
    return len(rows_to_move)
