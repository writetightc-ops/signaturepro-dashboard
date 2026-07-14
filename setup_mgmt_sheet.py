"""
Инициализация таблицы руководства SignaturePro (новая версия).

Создаёт 4 листа:
  - Ставки          : тарифные ставки + список сотрудников (читается дашбордом)
  - История_ЗП      : журнал зафиксированных выплат (пишется дашбордом)
  - Корректировки   : ручные корректировки ЗП (заполняется руководством вручную)
  - Итого_к_выплате : сводная таблица к выплате (формульная, для бухгалтерии)

Запуск:
  python setup_mgmt_sheet.py

Перед запуском:
  Добавь email сервисного аккаунта в настройки доступа таблицы руководства
  (Поделиться → Редактор). Email выводится при запуске.
"""

import time
import os
import sys

import gspread
from google.oauth2.service_account import Credentials

try:
    from config import CREDENTIALS_FILE
except ImportError:
    CREDENTIALS_FILE = "credentials.json"

MGMT_URL = "https://docs.google.com/spreadsheets/d/1bAjeDKCXtyp_MDlxJsnbyIFl43_l6JwbdXIkxmmpOhA/edit"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ─── Утилиты ──────────────────────────────────────────────────────────────────

def _rgb(r, g, b):
    return {"red": r / 255, "green": g / 255, "blue": b / 255}


def _with_retry(fn, max_retries=5):
    last_err = None
    for attempt in range(max_retries):
        try:
            return fn()
        except Exception as e:
            msg = str(e)
            if "429" in msg or "RESOURCE_EXHAUSTED" in msg or "quota" in msg.lower():
                wait = min(15 * (2 ** attempt), 120)
                print(f"    [API limit] жду {wait}с (попытка {attempt + 1}/{max_retries})...")
                time.sleep(wait)
                last_err = e
            else:
                raise
    raise RuntimeError(f"Превышен лимит API после {max_retries} попыток: {last_err}")


def get_client():
    creds_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), CREDENTIALS_FILE)
    if not os.path.exists(creds_path):
        raise FileNotFoundError(
            f"Файл учётных данных не найден: {creds_path}\n"
            "Скачай его из Google Cloud Console (Service Account → Keys → JSON)"
        )
    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    try:
        print(f"  Сервисный аккаунт: {creds.service_account_email}")
        print(f"  → Убедись, что этот email добавлен в новую таблицу как Редактор\n")
    except Exception:
        pass
    return gspread.authorize(creds)


# ─── Запросы форматирования ────────────────────────────────────────────────────

def _bold_row(ws_id, row_0, n_cols):
    return {"repeatCell": {
        "range": {"sheetId": ws_id, "startRowIndex": row_0, "endRowIndex": row_0 + 1,
                  "startColumnIndex": 0, "endColumnIndex": n_cols},
        "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
        "fields": "userEnteredFormat.textFormat.bold",
    }}


def _bg_row(ws_id, row_0, color, n_cols):
    return {"repeatCell": {
        "range": {"sheetId": ws_id, "startRowIndex": row_0, "endRowIndex": row_0 + 1,
                  "startColumnIndex": 0, "endColumnIndex": n_cols},
        "cell": {"userEnteredFormat": {"backgroundColor": color}},
        "fields": "userEnteredFormat.backgroundColor",
    }}


def _white_text_row(ws_id, row_0, n_cols):
    return {"repeatCell": {
        "range": {"sheetId": ws_id, "startRowIndex": row_0, "endRowIndex": row_0 + 1,
                  "startColumnIndex": 0, "endColumnIndex": n_cols},
        "cell": {"userEnteredFormat": {"textFormat": {"foregroundColor": _rgb(255, 255, 255), "bold": True}}},
        "fields": "userEnteredFormat.textFormat",
    }}


def _col_width(ws_id, col_start, col_end, px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": ws_id, "dimension": "COLUMNS",
                  "startIndex": col_start, "endIndex": col_end},
        "properties": {"pixelSize": px},
        "fields": "pixelSize",
    }}


def _freeze(ws_id, rows=1):
    return {"updateSheetProperties": {
        "properties": {"sheetId": ws_id, "gridProperties": {"frozenRowCount": rows}},
        "fields": "gridProperties.frozenRowCount",
    }}


# ═══════════════════════════ ЛИСТ: СТАВКИ ═════════════════════════════════════

TARIFF_RATES = [
    # Тариф | Варианты | Обучение | Бонус без правок | Тарифный бонус | Правка | Бонус менеджера
    ["E",            150, 200, 100,   0, 75, 200],
    ["ST",           150, 200, 100,   0, 75, 200],
    ["E_FAST",       150, 200, 150, 300, 75, 200],
    ["OPTNEW",       300, 400, 200,   0, 75, 300],
    ["PRNEW",        500, 200, 250, 300, 75, 300],
    ["ДОП_ОБУЧЕНИЕ",   0, 200,   0,   0,  0,   0],
    ["ДОП_ПОДПИСЬ",  250, 200, 100,   0, 75, 200],
]

EMPLOYEES = [
    # Имя | Роль | Коэффициент | Активен
    ["Марьям",          "Каллиграф", 1.0, "ДА"],
    ["Лена Вовина",     "Каллиграф", 1.0, "ДА"],
    ["Катерина Попова", "Каллиграф", 1.0, "ДА"],
    ["Катя Дорожкина",  "Каллиграф", 1.0, "ДА"],
    ["Ольга Струкова",  "Каллиграф", 1.0, "ДА"],
    ["Мария Тимофеева", "Каллиграф", 1.0, "ДА"],
]


def setup_rates_sheet(sh, ws):
    print("  Настраиваю лист 'Ставки'...")
    sid = ws.id

    values = (
        [["ТАРИФНЫЕ СТАВКИ"]]
        + [["Тариф", "Варианты", "Обучение", "Бонус без правок",
            "Тарифный бонус", "Правка", "Бонус менеджера"]]
        + TARIFF_RATES
        + [[]]
        + [["СОТРУДНИКИ"]]
        + [["Имя", "Роль", "Коэффициент", "Активен"]]
        + EMPLOYEES
        + [[]]
        + [["* Коэффициент — множитель ЗП каллиграфа (1.0 = стандарт, 1.2 = +20%). "
            "Активен = ДА/НЕТ. Изменения применяются в дашборде в течение 10 минут."]]
    )

    _with_retry(lambda: ws.clear())
    time.sleep(1)
    _with_retry(lambda: ws.update(range_name="A1", values=values, value_input_option="USER_ENTERED"))
    time.sleep(1)

    # Row indices (0-based): title=0, tariff_header=1, data=2..7, empty=8,
    # employees_title=9, emp_header=10, emp_data=11..16
    emp_title_row = 2 + len(TARIFF_RATES) + 1   # = 9
    emp_header_row = emp_title_row + 1            # = 10

    BLU_DARK  = _rgb(26, 115, 232)
    BLU_LIGHT = _rgb(197, 218, 255)
    GRN_DARK  = _rgb(52, 168, 83)
    GRN_LIGHT = _rgb(182, 215, 168)

    fmt = [
        _bg_row(sid, 0, BLU_DARK, 7),   _white_text_row(sid, 0, 7),   # ТАРИФНЫЕ СТАВКИ
        _bg_row(sid, 1, BLU_LIGHT, 7),  _bold_row(sid, 1, 7),          # col headers
        _bg_row(sid, emp_title_row, GRN_DARK, 4),  _white_text_row(sid, emp_title_row, 4),  # СОТРУДНИКИ
        _bg_row(sid, emp_header_row, GRN_LIGHT, 4), _bold_row(sid, emp_header_row, 4),       # col headers
        _freeze(sid, 2),
        _col_width(sid, 0, 1, 130),
        _col_width(sid, 1, 7, 145),
    ]
    _with_retry(lambda: sh.batch_update({"requests": fmt}))
    print("    [OK] Лист 'Ставки' готов")
    time.sleep(1)


# ═══════════════════════════ ЛИСТ: ИСТОРИЯ ЗП ═════════════════════════════════

HISTORY_HEADERS = [
    "Период", "Дата фиксации", "Сотрудник", "Роль",
    "Заказов (с начисл.)", "Завершено",
    "Варианты", "Правки", "Обучение", "Тарифный бонус",
    "Бонус без правок", "Бонус за заказ",
    "Итого (до коэф.)", "Коэффициент", "К выплате",
    "USA", "RU",
]


def setup_history_sheet(sh, ws):
    print("  Настраиваю лист 'История_ЗП'...")
    sid = ws.id

    _with_retry(lambda: ws.clear())
    time.sleep(1)
    _with_retry(lambda: ws.update(
        range_name="A1", values=[HISTORY_HEADERS], value_input_option="USER_ENTERED"
    ))
    time.sleep(1)

    nc = len(HISTORY_HEADERS)
    BLU = _rgb(26, 115, 232)

    fmt = [
        _bg_row(sid, 0, BLU, nc), _white_text_row(sid, 0, nc),
        _freeze(sid, 1),
        _col_width(sid, 0, 1, 210),
        _col_width(sid, 1, 2, 150),
        _col_width(sid, 2, 3, 170),
        _col_width(sid, 3, 4, 110),
        _col_width(sid, 4, 6, 120),
        _col_width(sid, 6, 15, 105),
        _col_width(sid, 15, 17, 90),
    ]
    _with_retry(lambda: sh.batch_update({"requests": fmt}))
    print("    [OK] Лист 'История_ЗП' готов")
    time.sleep(1)


# ═══════════════════════════ ЛИСТ: КОРРЕКТИРОВКИ ══════════════════════════════

ADJ_HEADERS = ["Период", "Сотрудник", "Тип", "Сумма", "Комментарий", "Дата добавления"]


def setup_adjustments_sheet(sh, ws):
    print("  Настраиваю лист 'Корректировки'...")
    sid = ws.id

    hint = (
        "Типы: Премия / Доплата / Аванс / Вычет / Штраф  |  "
        "Сумма: положительная = доплата, отрицательная = вычет  |  "
        "Формат периода: ДД.ММ.ГГГГ — ДД.ММ.ГГГГ"
    )

    _with_retry(lambda: ws.clear())
    time.sleep(1)
    _with_retry(lambda: ws.update(
        range_name="A1",
        values=[ADJ_HEADERS, [hint, "", "", "", "", ""]],
        value_input_option="USER_ENTERED"
    ))
    time.sleep(1)

    nc = len(ADJ_HEADERS)
    YLW = _rgb(251, 188, 4)

    fmt = [
        _bg_row(sid, 0, YLW, nc), _bold_row(sid, 0, nc),
        _freeze(sid, 1),
        _col_width(sid, 0, 1, 210),
        _col_width(sid, 1, 2, 170),
        _col_width(sid, 2, 3, 110),
        _col_width(sid, 3, 4, 100),
        _col_width(sid, 4, 5, 250),
        _col_width(sid, 5, 6, 140),
        # Dropdown for Тип column (C, index 2)
        {"setDataValidation": {
            "range": {"sheetId": sid, "startRowIndex": 2, "endRowIndex": 500,
                      "startColumnIndex": 2, "endColumnIndex": 3},
            "rule": {
                "condition": {"type": "ONE_OF_LIST", "values": [
                    {"userEnteredValue": "Премия"},
                    {"userEnteredValue": "Доплата"},
                    {"userEnteredValue": "Аванс"},
                    {"userEnteredValue": "Вычет"},
                    {"userEnteredValue": "Штраф"},
                ]},
                "showCustomUi": True, "strict": True,
            },
        }},
    ]
    _with_retry(lambda: sh.batch_update({"requests": fmt}))
    print("    [OK] Лист 'Корректировки' готов")
    time.sleep(1)


# ═══════════════════════════ ЛИСТ: ИТОГО К ВЫПЛАТЕ ════════════════════════════

def setup_total_sheet(sh, ws):
    print("  Настраиваю лист 'Итого_к_выплате'...")
    sid = ws.id

    _with_retry(lambda: ws.clear())
    time.sleep(1)

    # Static labels
    # Структура сводного блока:
    #   Row 1: Title
    #   Row 3: Период selector (B3/C3)
    #   Row 5: Итого к выплате (B5/C5)
    #   Row 6:   из них базовая ЗП (B6/C6)
    #   Row 7:   корректировки (B7/C7)
    #   Row 8:   USA (B8/C8)
    #   Row 9:   RU  (B9/C9)
    #   Row 11: Section header — ИСТОРИЯ ЗП | КОРРЕКТИРОВКИ
    #   Row 12: QUERY formulas
    labels = [
        ("A1",  [["ИТОГО К ВЫПЛАТЕ — SignaturePro ЗП Дашборд"]]),
        ("B3",  [["Период:"]]),
        ("C3",  [["(введи или выбери период из 'История_ЗП'!A:A)"]]),
        ("B5",  [["Итого к выплате (ЗП + корректировки):"]]),
        ("B6",  [["  из них базовая ЗП:"]]),
        ("B7",  [["  корректировки (\u00b1):"]]),
        ("B8",  [["  \U0001f1fa\U0001f1f8 из них USA:"]]),
        ("B9",  [["  \U0001f1f7\U0001f1fa из них RU:"]]),
        ("A11", [["ИСТОРИЯ ЗП ЗА ПЕРИОД"]]),
        ("J11", [["КОРРЕКТИРОВКИ ЗА ПЕРИОД"]]),
    ]
    # Formulas referencing C3
    formulas = [
        ("C5", [["=IFERROR(SUMIF('\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!A:A,C3,'\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!O:O)"
                 "+SUMIF('\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438'!A:A,C3,'\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438'!D:D),0)"]]),
        ("C6", [["=IFERROR(SUMIF('\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!A:A,C3,'\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!O:O),0)"]]),
        ("C7", [["=IFERROR(SUMIF('\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438'!A:A,C3,'\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438'!D:D),0)"]]),
        ("C8", [["=IFERROR(SUMIF('\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!A:A,C3,'\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!P:P),0)"]]),
        ("C9", [["=IFERROR(SUMIF('\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!A:A,C3,'\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!Q:Q),0)"]]),
        ("A12", [["=IFERROR(QUERY('\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!A:Q,"
                  "\"SELECT C,D,F,O,P,Q WHERE A='\"&C3&\"' "
                  "ORDER BY D,C "
                  "LABEL C '\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a',D '\u0420\u043e\u043b\u044c',F '\u0417\u0430\u0432\u0435\u0440\u0448\u0435\u043d\u043e',"
                  "O '\u041a \u0432\u044b\u043f\u043b\u0430\u0442\u0435',P 'USA',Q 'RU'\",0),"
                  "\"\u041d\u0435\u0442 \u0434\u0430\u043d\u043d\u044b\u0445 \u0437\u0430 \u0432\u044b\u0431\u0440\u0430\u043d\u043d\u044b\u0439 \u043f\u0435\u0440\u0438\u043e\u0434\")"]]),
        ("J12", [["=IFERROR(QUERY('\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438'!A:F,"
                  "\"SELECT B,C,D,E WHERE A='\"&C3&\"' "
                  "ORDER BY B "
                  "LABEL B '\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a',C '\u0422\u0438\u043f',D '\u0421\u0443\u043c\u043c\u0430',E '\u041a\u043e\u043c\u043c\u0435\u043d\u0442\u0430\u0440\u0438\u0439'\",0),"
                  "\"\u041d\u0435\u0442 \u043a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043e\u043a\")"]]),
    ]

    for cell, val in labels + formulas:
        _with_retry(lambda c=cell, v=val: ws.update(
            range_name=c, values=v, value_input_option="USER_ENTERED"
        ))
        time.sleep(0.4)

    BLU     = _rgb(26, 115, 232)
    GRY     = _rgb(241, 241, 241)
    BLU_LT  = _rgb(219, 234, 254)   # USA row bg
    ROSE_LT = _rgb(254, 226, 226)   # RU row bg

    fmt = [
        _bg_row(sid, 0, BLU, 14), _white_text_row(sid, 0, 14),  # title row
        _bold_row(sid, 4, 5),                                     # Итого к выплате (row 5=idx 4)
        _bg_row(sid, 10, GRY, 14), _bold_row(sid, 10, 14),       # section headers (row 11=idx 10)
        # Подсветка USA/RU строк (rows 8-9 = idx 7-8)
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 7, "endRowIndex": 8,
                      "startColumnIndex": 1, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": {"backgroundColor": BLU_LT}},
            "fields": "userEnteredFormat.backgroundColor",
        }},
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 8, "endRowIndex": 9,
                      "startColumnIndex": 1, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": {"backgroundColor": ROSE_LT}},
            "fields": "userEnteredFormat.backgroundColor",
        }},
        # Data validation dropdown for C3 (period selector)
        {"setDataValidation": {
            "range": {"sheetId": sid, "startRowIndex": 2, "endRowIndex": 3,
                      "startColumnIndex": 2, "endColumnIndex": 3},
            "rule": {
                "condition": {"type": "ONE_OF_RANGE",
                              "values": [{"userEnteredValue": "='\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!$A$2:$A$1000"}]},
                "showCustomUi": True, "strict": False,
            },
        }},
        _col_width(sid, 1, 2, 290),
        _col_width(sid, 2, 3, 220),
        _col_width(sid, 0, 1, 170),
        # Currency format для C5:C9 (Итого + компоненты + USA + RU)
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 4, "endRowIndex": 9,
                      "startColumnIndex": 2, "endColumnIndex": 3},
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "CURRENCY", "pattern": "# ##0 \"\u20bd\""},
                "textFormat": {"bold": True},
            }},
            "fields": "userEnteredFormat(numberFormat,textFormat)",
        }},
        # Italic для строк USA/RU (idx 7-8)
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 7, "endRowIndex": 9,
                      "startColumnIndex": 2, "endColumnIndex": 3},
            "cell": {"userEnteredFormat": {"textFormat": {"italic": True, "bold": False}}},
            "fields": "userEnteredFormat.textFormat",
        }},
    ]
    _with_retry(lambda: sh.batch_update({"requests": fmt}))
    print("    [OK] Лист 'Итого_к_выплате' готов")
    time.sleep(1)


# ═══════════════════════════ MAIN ═════════════════════════════════════════════

def main():
    print("=" * 60)
    print("  SignaturePro — Инициализация таблицы руководства (NEW)")
    print("=" * 60)

    creds_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), CREDENTIALS_FILE)
    if not os.path.exists(creds_path):
        print(f"\n  Ошибка: файл '{CREDENTIALS_FILE}' не найден рядом со скриптом.")
        sys.exit(1)

    print("\nПодключение к Google Sheets API...")
    gc = get_client()

    print(f"Открываю таблицу руководства...")
    try:
        sh = _with_retry(lambda: gc.open_by_url(MGMT_URL))
    except Exception as e:
        print(f"\n  Ошибка доступа к таблице: {e}")
        print("  Добавь email сервисного аккаунта (см. выше) в таблицу как Редактор и повтори.")
        sys.exit(1)

    print(f"  Открыта: {sh.title}")
    existing = {ws.title for ws in sh.worksheets()}
    print(f"  Существующие листы: {sorted(existing)}\n")

    def get_or_create(title, rows=500, cols=20):
        if title in existing:
            print(f"  Лист '{title}' уже существует — перезаписываю...")
            return sh.worksheet(title)
        print(f"  Создаю лист '{title}'...")
        ws = _with_retry(lambda: sh.add_worksheet(title=title, rows=rows, cols=cols))
        time.sleep(1)
        return ws

    ws_rates = get_or_create("Ставки", rows=50, cols=10)
    setup_rates_sheet(sh, ws_rates)

    ws_hist = get_or_create("История_ЗП", rows=2000, cols=18)
    setup_history_sheet(sh, ws_hist)

    ws_adj = get_or_create("Корректировки", rows=500, cols=7)
    setup_adjustments_sheet(sh, ws_adj)

    ws_total = get_or_create("Итого_к_выплате", rows=200, cols=15)
    setup_total_sheet(sh, ws_total)

    # Удалить Sheet1 если это единственный лишний лист
    for dummy in ("Sheet1", "Лист1"):
        if dummy in existing and len(sh.worksheets()) > 4:
            try:
                _with_retry(lambda d=dummy: sh.del_worksheet(sh.worksheet(d)))
                print(f"  Удалён пустой лист '{dummy}'")
                time.sleep(1)
            except Exception:
                pass

    print("\n" + "=" * 60)
    print("  [DONE] Таблица руководства инициализирована!")
    print(f"  URL: {MGMT_URL}")
    print("  Листы: Ставки | История_ЗП | Корректировки | Итого_к_выплате")
    print("\n  Следующий шаг: запусти dashboard.py — он подхватит новую таблицу.")
    print("=" * 60)


if __name__ == "__main__":
    main()
