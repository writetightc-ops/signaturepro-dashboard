"""
Адаптер для чтения/записи данных из Google Таблицы через Google Sheets API.
Использует gspread + google-auth (сервисный аккаунт).
"""

import time
import os
import re as _re
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

def _parse_date(val, year=None):
    """
    Преобразует строку или datetime/date объект в date.
    Поддерживает полные форматы (dd.mm.yyyy) и короткие (d.MM / dd.MM без года).
    Для коротких форматов используется year (год из названия листа) или текущий год.
    """
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
    # Полные форматы с годом
    for fmt in ["%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d.%m.%y",
                "%m/%d/%Y", "%d-%m-%Y", "%d-%m-%y"]:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    # Короткий формат: d.MM или dd.MM (без года) — используем год листа
    m = _re.match(r'^(\d{1,2})\.(\d{1,2})$', s)
    if m:
        y = year or date.today().year
        try:
            return date(y, int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
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
    """
    Преобразует строку листа (список значений) в словарь заказа.
    Год для коротких дат (d.MM) берётся из названия листа, напр. "Апрель_2026" → 2026.
    """
    r = list(row) + [""] * 25  # на случай если строка короткая
    # Пустая строка — вернём None
    if not r[0] and not r[2]:
        return None
    # Строка-подсказка из шаблона
    if isinstance(r[0], str) and r[0].startswith("←"):
        return None

    # Извлекаем год из имени листа (напр. "Апрель_2026" → 2026)
    year_match = _re.search(r'(\d{4})', sheet_name)
    sheet_year = int(year_match.group(1)) if year_match else date.today().year

    def pd(val):
        return _parse_date(val, sheet_year)

    # Если "варианты" заполнен, но не является датой (например, галочка, цифра, "✓"),
    # считаем что варианты были сданы в дату поступления заказа (fallback).
    # Это нужно для корректного подсчёта ЗП в дашборде по старым данным без дат.
    _variants_raw = str(r[6]).strip() if r[6] else ""
    _variants_date = pd(r[6])
    if _variants_date is None and _variants_raw:
        _variants_date = pd(r[0])  # fallback: поступление

    return {
        "поступление": pd(r[0]),
        "статус":      _parse_bool(r[1]),
        "менеджер":    str(r[2]).strip() if r[2] else "",
        "каллиграф":   str(r[3]).strip() if r[3] else "",
        "клиент":      str(r[4]).strip() if r[4] else "",
        "тариф":       str(r[5]).strip().upper() if r[5] else "",
        "варианты":    _variants_date,
        "ос1":         pd(r[7]),
        "правка1":     pd(r[8]),
        "ос2":         pd(r[9]),
        "правка2":     pd(r[10]),
        "ос3":         pd(r[11]),
        "правка3":     pd(r[12]),
        "ос4":         pd(r[13]),
        "правка4":     pd(r[14]),
        "обучение":    pd(r[15]),
        "ссылка":      str(r[16]).strip() if r[16] else "",
        "завершение":  pd(r[17]),
        "заметки":     str(r[18]).strip() if len(r) > 18 and r[18] else "",
        # Столбцы T–W: ОС 5, Правка 5, ОС 6, Правка 6
        "ос5":         pd(r[19]),
        "правка5":     pd(r[20]),
        "ос6":         pd(r[21]),
        "правка6":     pd(r[22]),
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


# ─── Таблица руководства: ставки, история ЗП, корректировки ──────────────────

_mgmt_cache = {}  # {key: {"ts": float, ...data}}


def _mgmt_key(url, suffix):
    return url + "::" + suffix


def _mgmt_ws(mgmt_url, credentials_file, sheet_title):
    """Открывает лист таблицы руководства."""
    client = _get_client(credentials_file)
    sh = _with_retry(lambda: client.open_by_url(mgmt_url))
    try:
        return sh, sh.worksheet(sheet_title)
    except Exception:
        raise ValueError(
            f"Лист '{sheet_title}' не найден в таблице руководства. "
            f"Запусти setup_mgmt_sheet.py для инициализации."
        )


def read_rates(mgmt_url, credentials_file, cache_ttl=600):
    """
    Читает тарифные ставки и коэффициенты сотрудников из листа 'Ставки'.
    Возвращает (calligrapher_rates, manager_rates, employees).
    Кешируется на cache_ttl секунд (по умолчанию 10 минут).
    """
    key = _mgmt_key(mgmt_url, "rates")
    now = time.time()
    cached = _mgmt_cache.get(key)
    if cached and (now - cached["ts"]) < cache_ttl:
        return cached["cal"], cached["mgr"], cached["emp"]

    _, ws = _mgmt_ws(mgmt_url, credentials_file, "Ставки")
    rows = _with_retry(ws.get_all_values)

    calligrapher_rates = {}
    manager_rates = {}
    employees = {}
    section = None  # 'tariff' | 'employee' | None

    def _int(cell):
        try:
            return int(float(str(cell).replace(",", ".").strip() or "0"))
        except (ValueError, TypeError):
            return 0

    def _float(cell):
        try:
            return float(str(cell).replace(",", ".").strip() or "1.0")
        except (ValueError, TypeError):
            return 1.0

    for row in rows:
        if not row or not any(row):
            section = None
            continue
        first = str(row[0]).strip()
        if not first:
            continue

        if first == "Тариф":
            section = "tariff"
            continue
        if first == "Имя":
            section = "employee"
            continue
        if first in ("ТАРИФНЫЕ СТАВКИ", "СОТРУДНИКИ") or first.startswith("*"):
            continue

        if section == "tariff":
            tariff = first.upper()
            r = row + [""] * 8
            calligrapher_rates[tariff] = {
                "варианты":         _int(r[1]),
                "обучение":         _int(r[2]),
                "бонус_без_правок": _int(r[3]),
                "тарифный_бонус":   _int(r[4]),
                "правка":           _int(r[5]),
            }
            manager_rates[tariff] = _int(r[6])

        elif section == "employee":
            r = row + [""] * 5
            name = first
            role   = str(r[1]).strip()
            coeff  = _float(r[2])
            active = str(r[3]).strip().upper() in ("ДА", "YES", "TRUE", "1", "+", "ИСТИНА")
            if active:
                employees[name] = {"role": role, "coefficient": coeff}

    _mgmt_cache[key] = {"ts": now, "cal": calligrapher_rates, "mgr": manager_rates, "emp": employees}
    return calligrapher_rates, manager_rates, employees


def read_closed_periods(mgmt_url, credentials_file, cache_ttl=30):
    """Возвращает список зафиксированных периодов из 'История_ЗП'."""
    key = _mgmt_key(mgmt_url, "closed")
    now = time.time()
    cached = _mgmt_cache.get(key)
    if cached and (now - cached["ts"]) < cache_ttl:
        return cached["periods"]

    _, ws = _mgmt_ws(mgmt_url, credentials_file, "История_ЗП")
    rows = _with_retry(ws.get_all_values)

    periods = list(dict.fromkeys(
        row[0].strip() for row in rows[1:] if row and row[0].strip()
    ))
    _mgmt_cache[key] = {"ts": now, "periods": periods}
    return periods


def write_salary_history(mgmt_url, credentials_file, date_from, date_to, employees_data):
    """
    Фиксирует ЗП за период в листе 'История_ЗП'.
    Возвращает количество записанных строк.
    Поднимает ValueError если период уже зафиксирован.
    """
    period_str = f"{date_from.strftime('%d.%m.%Y')} \u2014 {date_to.strftime('%d.%m.%Y')}"
    now_str    = date.today().strftime("%d.%m.%Y")

    # Проверяем: период уже закрыт?
    closed = read_closed_periods(mgmt_url, credentials_file, cache_ttl=5)
    if period_str in closed:
        raise ValueError(f"Период «{period_str}» уже зафиксирован в истории")

    _, ws = _mgmt_ws(mgmt_url, credentials_file, "История_ЗП")

    rows_to_add = []
    for emp in employees_data.values():
        bd    = emp.get("breakdown", {})
        coeff = emp.get("coefficient", 1.0)
        total = emp.get("total", 0.0)
        base  = round(total / coeff, 2) if coeff and coeff != 0 else total

        rows_to_add.append([
            period_str,
            now_str,
            emp.get("name", ""),
            emp.get("role", ""),
            emp.get("orders_count", 0),
            emp.get("orders_done", 0),
            round(bd.get("варианты",         0) / coeff, 2) if coeff else bd.get("варианты",         0),
            round(bd.get("правки",           0) / coeff, 2) if coeff else bd.get("правки",           0),
            round(bd.get("обучение",         0) / coeff, 2) if coeff else bd.get("обучение",         0),
            round(bd.get("тарифный_бонус",   0) / coeff, 2) if coeff else bd.get("тарифный_бонус",   0),
            round(bd.get("бонус_без_правок", 0) / coeff, 2) if coeff else bd.get("бонус_без_правок", 0),
            round(bd.get("бонус_за_заказ",   0),         2),
            base,
            coeff,
            total,
            round(emp.get("total_usa", 0), 2),
            round(emp.get("total_ru",  0), 2),
        ])

    if not rows_to_add:
        raise ValueError("Нет начислений за выбранный период — нечего фиксировать")

    _with_retry(lambda: ws.append_rows(rows_to_add, value_input_option="USER_ENTERED"))

    # Сбрасываем кеш закрытых периодов
    key = _mgmt_key(mgmt_url, "closed")
    if key in _mgmt_cache:
        del _mgmt_cache[key]

    return len(rows_to_add)


def invalidate_mgmt_cache(mgmt_url):
    """Сбрасывает весь кеш таблицы руководства (ставки, история, корректировки)."""
    for suffix in ("rates", "closed"):
        key = _mgmt_key(mgmt_url, suffix)
        if key in _mgmt_cache:
            del _mgmt_cache[key]


# ─── Форматирование листа (условное + валидация дат) ─────────────────────────

# 0-based индексы столбцов с датами (A=0, G=6, H=7, ...)
# G(6)=варианты включён → date validation + формат dd.mm
_DATE_COL_INDICES = [0, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 17, 19, 20, 21, 22]
_MAX_ROWS = 500
_N_COLS   = 24   # A–X


def _rgb(r, g, b):
    return {"red": r / 255, "green": g / 255, "blue": b / 255}


def _build_total_formula(r):
    """Строит формулу «Итого ЗП» для строки r (1-based номер строки листа)."""
    var = (
        f'IF(G{r}<>"",'
        f'IF(OR(F{r}="E",F{r}="ST",F{r}="E_FAST"),150,'
        f'IF(OR(F{r}="OPTNEW",F{r}="PRNEW"),500,'
        f'IF(F{r}="OPT2",300,0))),0)'
    )
    pravki    = f'COUNTA(I{r},K{r},M{r},O{r},U{r},W{r})*75'
    obuchen   = f'IF(P{r}<>"",200,0)'
    tar_bonus = f'IF(AND(G{r}<>"",OR(F{r}="E_FAST",F{r}="PRNEW")),300,0)'
    no_pravki = f'COUNTA(I{r},K{r},M{r},O{r},U{r},W{r})=0'
    bonus_bp  = (
        f'IF(AND(B{r}=TRUE,G{r}<>"",{no_pravki}),'
        f'IF(OR(F{r}="E",F{r}="ST"),100,'
        f'IF(F{r}="E_FAST",150,'
        f'IF(F{r}="OPTNEW",200,'
        f'IF(F{r}="PRNEW",250,'
        f'IF(F{r}="OPT2",150,0))))),0)'
    )
    return f'={var}+{pravki}+{obuchen}+{tar_bonus}+{bonus_bp}'


def _apply_total_formula(ws):
    """Записывает формулу «Итого ЗП» в столбец X для всех строк с данными."""
    all_vals = _with_retry(ws.get_all_values)
    data_rows = [
        r for r in all_vals[1:]
        if (len(r) > 0 and r[0].strip()) or (len(r) > 2 and r[2].strip())
    ]
    last_row = 1 + len(data_rows)
    if last_row < 2:
        return

    formulas = [[_build_total_formula(r)] for r in range(2, last_row + 1)]
    _with_retry(lambda: ws.update(
        range_name=f"X2:X{last_row}",
        values=formulas,
        value_input_option="USER_ENTERED",
    ))
    time.sleep(1)


def _apply_sheet_formatting(sh, ws):
    """
    Применяет к листу:
      1. Три правила условного форматирования (зелёный / жёлтый / серый)
      2. Data validation (календарь) на все столбцы с датами
      3. Формула «Итого ЗП» в столбце X
    Вызывается при создании нового листа и после каждого переноса.
    """
    sid = ws.id

    # Сначала удаляем старые правила условного форматирования листа
    ws_meta = _with_retry(lambda: sh.fetch_sheet_metadata())
    for sheet_data in ws_meta.get("sheets", []):
        if sheet_data["properties"]["sheetId"] == sid:
            n_rules = len(sheet_data.get("conditionalFormats", []))
            break
    else:
        n_rules = 0

    del_requests = [
        {"deleteConditionalFormatRule": {"sheetId": sid, "index": 0}}
        for _ in range(n_rules)
    ]
    if del_requests:
        _with_retry(lambda r=del_requests: sh.batch_update({"requests": r}))
        time.sleep(1)

    grid = {
        "sheetId": sid,
        "startRowIndex": 1,
        "endRowIndex": _MAX_ROWS + 1,
        "startColumnIndex": 0,
        "endColumnIndex": _N_COLS,
    }

    fmt_requests = [
        # Статус TRUE → зелёный
        {"addConditionalFormatRule": {"rule": {"ranges": [grid], "booleanRule": {
            "condition": {"type": "CUSTOM_FORMULA",
                          "values": [{"userEnteredValue": "=$B2=TRUE"}]},
            "format": {"backgroundColor": _rgb(198, 239, 206)},
        }}, "index": 0}},
        # E / ST / E_FAST → светло-жёлтый
        {"addConditionalFormatRule": {"rule": {"ranges": [grid], "booleanRule": {
            "condition": {"type": "CUSTOM_FORMULA",
                          "values": [{"userEnteredValue": '=OR($F2="E",$F2="ST",$F2="E_FAST")'}]},
            "format": {"backgroundColor": _rgb(255, 255, 204)},
        }}, "index": 1}},
        # OPTNEW / OPT2 / PRNEW → светло-серый
        {"addConditionalFormatRule": {"rule": {"ranges": [grid], "booleanRule": {
            "condition": {"type": "CUSTOM_FORMULA",
                          "values": [{"userEnteredValue": '=OR($F2="OPTNEW",$F2="OPT2",$F2="PRNEW")'}]},
            "format": {"backgroundColor": _rgb(240, 240, 240)},
        }}, "index": 2}},
    ]
    _with_retry(lambda: sh.batch_update({"requests": fmt_requests}))
    time.sleep(1)

    # Data validation (календарь) + формат dd.mm на столбцы с датами
    val_requests = []
    for col_idx in _DATE_COL_INDICES:
        date_range = {
            "sheetId": sid,
            "startRowIndex": 1,
            "endRowIndex": _MAX_ROWS + 1,
            "startColumnIndex": col_idx,
            "endColumnIndex": col_idx + 1,
        }
        val_requests.append({
            "setDataValidation": {
                "range": date_range,
                "rule": {
                    "condition": {"type": "DATE_IS_VALID"},
                    "showCustomUi": True,
                    "strict": False,
                },
            }
        })
        val_requests.append({
            "repeatCell": {
                "range": date_range,
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {"type": "DATE", "pattern": "dd.mm"}
                    }
                },
                "fields": "userEnteredFormat.numberFormat",
            }
        })
    for i in range(0, len(val_requests), 10):
        batch = val_requests[i:i + 10]
        _with_retry(lambda b=batch: sh.batch_update({"requests": b}))
        time.sleep(1)

    # Чекбокс (столбец B — «Статус»)
    _with_retry(lambda: sh.batch_update({"requests": [{
        "setDataValidation": {
            "range": {
                "sheetId": sid,
                "startRowIndex": 1,
                "endRowIndex": _MAX_ROWS + 1,
                "startColumnIndex": 1,
                "endColumnIndex": 2,
            },
            "rule": {
                "condition": {"type": "BOOLEAN"},
                "showCustomUi": True,
                "strict": True,
            },
        }
    }]}))
    time.sleep(1)

    # Формула «Итого ЗП» в столбце X
    _apply_total_formula(ws)


def transfer_orders(spreadsheet_url, credentials_file, from_sheet_name, to_sheet_name):
    """
    Переносит незавершённые заказы (без даты завершения) из одного листа в другой.
    Копирует строки через copyPaste (сохраняет форматирование ячеек),
    затем применяет условное форматирование и валидацию дат к целевому листу.
    Возвращает количество перенесённых строк.
    """
    client = _get_client(credentials_file)
    sh = _with_retry(lambda: client.open_by_url(spreadsheet_url))

    # Исходный лист
    try:
        ws_from = sh.worksheet(from_sheet_name)
    except Exception:
        raise ValueError(f"Лист «{from_sheet_name}» не найден в таблице")

    all_rows = _with_retry(ws_from.get_all_values)
    if len(all_rows) < 2:
        return 0

    headers   = all_rows[0]
    data_rows = all_rows[1:]

    DONE_COL = 17  # 0-based: «Дата завершения»

    # Определяем строки для переноса (0-based в листе, строка 0 = заголовок)
    move_row_indices = []  # 0-based индексы строк листа (заголовок = 0)
    rows_to_keep = [headers]

    for i, row in enumerate(data_rows):
        r = row + [""] * 5
        if not r[0] and not r[2]:
            continue
        done = r[DONE_COL].strip() if DONE_COL < len(r) else ""
        if not done:
            move_row_indices.append(i + 1)  # +1 — заголовок занимает строку 0
        else:
            rows_to_keep.append(row)

    if not move_row_indices:
        return 0

    # Целевой лист — создаём если нет
    is_new_sheet = False
    try:
        ws_to = sh.worksheet(to_sheet_name)
    except Exception:
        ws_to = _with_retry(lambda: sh.add_worksheet(
            title=to_sheet_name, rows=600, cols=_N_COLS + 2))
        _with_retry(lambda: ws_to.append_row(headers, value_input_option="USER_ENTERED"))
        time.sleep(1)
        is_new_sheet = True

    # Текущий последний ряд в целевом листе (0-based: куда вставлять)
    dest_all = _with_retry(ws_to.get_all_values)
    dest_start = len(dest_all)  # следующая строка после последней

    from_id = ws_from.id
    to_id   = ws_to.id

    # Копируем строки через copyPaste (сохраняет форматирование ячеек и валидацию)
    copy_requests = []
    for offset, src_row_idx in enumerate(move_row_indices):
        copy_requests.append({
            "copyPaste": {
                "source": {
                    "sheetId": from_id,
                    "startRowIndex": src_row_idx,
                    "endRowIndex": src_row_idx + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": _N_COLS,
                },
                "destination": {
                    "sheetId": to_id,
                    "startRowIndex": dest_start + offset,
                    "endRowIndex": dest_start + offset + 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": _N_COLS,
                },
                "pasteType": "PASTE_NORMAL",
                "pasteOrientation": "NORMAL",
            }
        })

    BATCH = 50
    for i in range(0, len(copy_requests), BATCH):
        batch = copy_requests[i:i + BATCH]
        _with_retry(lambda b=batch: sh.batch_update({"requests": b}))
        time.sleep(1)

    # Применяем условное форматирование и валидацию к целевому листу
    _apply_sheet_formatting(sh, ws_to)
    time.sleep(1)

    # Удаляем перенесённые строки из исходного листа (сохраняем форматирование оставшихся)
    delete_requests = []
    for row_idx in sorted(move_row_indices, reverse=True):
        delete_requests.append({
            "deleteDimension": {
                "range": {
                    "sheetId": from_id,
                    "dimension": "ROWS",
                    "startIndex": row_idx,
                    "endIndex": row_idx + 1,
                }
            }
        })
    if delete_requests:
        _with_retry(lambda: sh.batch_update({"requests": delete_requests}))
        time.sleep(1)
        # Восстанавливаем формулу «Итого ЗП» в исходном листе
        _apply_total_formula(ws_from)

    invalidate_cache(spreadsheet_url)
    return len(move_row_indices)
