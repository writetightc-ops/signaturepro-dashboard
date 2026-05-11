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
    ["E",      150, 200, 100,   0, 75, 200],
    ["ST",     150, 200, 100,   0, 75, 200],
    ["E_FAST", 150, 200, 150, 300, 75, 200],
    ["OPTNEW", 500, 200, 200,   0, 75, 300],
    ["PRNEW",  500, 200, 250, 300, 75, 300],
    ["OPT2",   300, 200, 150,   0, 75, 300],
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
    """
    Лист Корректировки — выпадающие списки:
      A (Период)     : из История_ЗП!$A$2:$A$1000
      B (Сотрудник)  : из Ставки!$A$12:$A$100
      C (Тип)        : Премия / Доплата / Аванс / Вычет / Штраф
      D (Сумма)      : денежный формат (+ доплата, − вычет)
      E (Комментарий): текст
      F (Дата)       : дата добавления
    """
    print("  Настраиваю лист 'Корректировки'...")
    sid = ws.id

    _with_retry(lambda: ws.clear())
    time.sleep(1)
    _with_retry(lambda: ws.update(
        range_name="A1",
        values=[ADJ_HEADERS],
        value_input_option="USER_ENTERED"
    ))
    time.sleep(1)

    nc = len(ADJ_HEADERS)
    YLW = _rgb(251, 188, 4)

    fmt = [
        _bg_row(sid, 0, YLW, nc), _bold_row(sid, 0, nc),
        _freeze(sid, 1),
        _col_width(sid, 0, 1, 220),
        _col_width(sid, 1, 2, 180),
        _col_width(sid, 2, 3, 120),
        _col_width(sid, 3, 4, 110),
        _col_width(sid, 4, 5, 260),
        _col_width(sid, 5, 6, 150),

        # A (Период): выпадающий список из История_ЗП
        {"setDataValidation": {
            "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": 500,
                      "startColumnIndex": 0, "endColumnIndex": 1},
            "rule": {
                "condition": {
                    "type": "ONE_OF_RANGE",
                    "values": [{"userEnteredValue": "='\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!$A$2:$A$1000"}]
                },
                "showCustomUi": True, "strict": False,
            },
        }},

        # B (Сотрудник): выпадающий список из Ставки
        {"setDataValidation": {
            "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": 500,
                      "startColumnIndex": 1, "endColumnIndex": 2},
            "rule": {
                "condition": {
                    "type": "ONE_OF_RANGE",
                    "values": [{"userEnteredValue": "='\u0421\u0442\u0430\u0432\u043a\u0438'!$A$12:$A$100"}]
                },
                "showCustomUi": True, "strict": False,
            },
        }},

        # C (Тип): фиксированный список
        {"setDataValidation": {
            "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": 500,
                      "startColumnIndex": 2, "endColumnIndex": 3},
            "rule": {
                "condition": {"type": "ONE_OF_LIST", "values": [
                    {"userEnteredValue": "\u041f\u0440\u0435\u043c\u0438\u044f"},
                    {"userEnteredValue": "\u0414\u043e\u043f\u043b\u0430\u0442\u0430"},
                    {"userEnteredValue": "\u0410\u0432\u0430\u043d\u0441"},
                    {"userEnteredValue": "\u0412\u044b\u0447\u0435\u0442"},
                    {"userEnteredValue": "\u0428\u0442\u0440\u0430\u0444"},
                ]},
                "showCustomUi": True, "strict": True,
            },
        }},

        # D (Сумма): денежный формат
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": 500,
                      "startColumnIndex": 3, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "CURRENCY", "pattern": "# ##0 \"\u20bd\""}
            }},
            "fields": "userEnteredFormat.numberFormat",
        }},
    ]
    _with_retry(lambda: sh.batch_update({"requests": fmt}))
    print("    [OK] Лист 'Корректировки' готов (дропдауны Период, Сотрудник, Тип)")
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
        ("B5",  [["Итого к выплате (ЗП + доплаты):"]]),
        ("B6",  [["  из них базовая ЗП:"]]),
        ("B7",  [["  доплаты / вычеты (\u00b1):"]]),
        ("B8",  [["  \U0001f1fa\U0001f1f8 из них USA:"]]),
        ("B9",  [["  \U0001f1f7\U0001f1fa из них RU:"]]),
        # Раздел: сводка по сотрудникам (строки 11-12)
        ("A11", [["\u0421\u0412\u041e\u0414\u041a\u0410 \u041f\u041e \u0421\u041e\u0422\u0420\u0423\u0414\u041d\u0418\u041a\u0410\u041c \u2014 \u0411\u0410\u0417\u041e\u0412\u0410\u042f \u0417\u041f + \u0414\u041e\u041f\u041b\u0410\u0422\u042b"]]),
        ("A12", [["\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a",
                  "\u0411\u0430\u0437\u043e\u0432\u0430\u044f \u0417\u041f",
                  "\u0414\u043e\u043f\u043b\u0430\u0442\u044b (\u00b1)",
                  "\u0418\u0442\u043e\u0433\u043e \u043a \u0432\u044b\u043f\u043b\u0430\u0442\u0435"]]),
        # Разделители истории и доплат (строки 29-30)
        ("A29", [["\u0418\u0421\u0422\u041e\u0420\u0418\u042f \u0417\u041f \u0417\u0410 \u041f\u0415\u0420\u0418\u041e\u0414"]]),
        ("H29", [["\u0414\u041e\u041f\u041b\u0410\u0422\u042b \u0417\u0410 \u041f\u0415\u0420\u0418\u041e\u0414"]]),
    ]

    # Формулы сводных сумм (C5-C9)
    # Формулы SUMIFS по сотрудникам (B13:D27)
    per_emp_formulas = []
    for r in range(13, 28):
        b = (
            f"=IF(A{r}=\"\",\"\","
            f"IFERROR(SUMIFS('\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!O:O,"
            f"'\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!A:A,$C$3,"
            f"'\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!C:C,A{r}),0))"
        )
        c = (
            f"=IF(A{r}=\"\",\"\","
            f"IFERROR(SUMIFS('\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438'!D:D,"
            f"'\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438'!A:A,$C$3,"
            f"'\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438'!B:B,A{r}),0))"
        )
        d = f"=IF(A{r}=\"\",\"\",B{r}+C{r})"
        per_emp_formulas.append((f"B{r}:D{r}", [[b, c, d]]))

    formulas = [
        ("C5", [["=IFERROR(SUMIF('\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!A:A,C3,'\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!O:O)"
                 "+SUMIF('\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438'!A:A,C3,'\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438'!D:D),0)"]]),
        ("C6", [["=IFERROR(SUMIF('\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!A:A,C3,'\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!O:O),0)"]]),
        ("C7", [["=IFERROR(SUMIF('\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438'!A:A,C3,'\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438'!D:D),0)"]]),
        ("C8", [["=IFERROR(SUMIF('\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!A:A,C3,'\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!P:P),0)"]]),
        ("C9", [["=IFERROR(SUMIF('\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!A:A,C3,'\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!Q:Q),0)"]]),
        # QUERY список сотрудников (A13, спускается вниз)
        ("A13", [["=IFERROR(QUERY('\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!A:C,"
                  "\"SELECT C WHERE A='\"&$C$3&\"' GROUP BY C ORDER BY C\",0),\"\")"]]),
        # История ЗП (A30)
        ("A30", [["=IFERROR(QUERY('\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!A:Q,"
                  "\"SELECT C,D,F,O,P,Q WHERE A='\"&C3&\"' "
                  "ORDER BY D,C "
                  "LABEL C '\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a',D '\u0420\u043e\u043b\u044c',F '\u0417\u0430\u0432\u0435\u0440\u0448\u0435\u043d\u043e',"
                  "O '\u041a \u0432\u044b\u043f\u043b\u0430\u0442\u0435',P 'USA',Q 'RU'\",0),"
                  "\"\u041d\u0435\u0442 \u0434\u0430\u043d\u043d\u044b\u0445\")"]]),
        # Доплаты (H30)
        ("H30", [["=IFERROR(QUERY('\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438'!A:F,"
                  "\"SELECT B,C,D,E WHERE A='\"&C3&\"' "
                  "ORDER BY B "
                  "LABEL B '\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a',C '\u0422\u0438\u043f',D '\u0421\u0443\u043c\u043c\u0430',E '\u041a\u043e\u043c\u043c\u0435\u043d\u0442\u0430\u0440\u0438\u0439'\",0),"
                  "\"\u041d\u0435\u0442 \u0434\u043e\u043f\u043b\u0430\u0442\")"]]),
    ] + per_emp_formulas

    for cell, val in labels + formulas:
        _with_retry(lambda c=cell, v=val: ws.update(
            range_name=c, values=v, value_input_option="USER_ENTERED"
        ))
        time.sleep(0.4)

    BLU     = _rgb(26, 115, 232)
    GRY     = _rgb(241, 241, 241)
    GRN_DRK = _rgb(52, 168, 83)
    GRN_LT  = _rgb(182, 215, 168)
    BLU_LT  = _rgb(219, 234, 254)   # USA row bg
    ROSE_LT = _rgb(254, 226, 226)   # RU row bg

    fmt = [
        _bg_row(sid, 0, BLU, 14), _white_text_row(sid, 0, 14),  # title row
        _bold_row(sid, 4, 5),                                     # Итого к выплате (row 5=idx 4)
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
        # Сводка по сотрудникам: заголовок раздела (row 11 = idx 10)
        _bg_row(sid, 10, GRN_DRK, 5), _white_text_row(sid, 10, 5),
        # Хедеры колонок (row 12 = idx 11)
        _bg_row(sid, 11, GRN_LT, 5), _bold_row(sid, 11, 5),
        # Заголовки истории/доплат (row 29 = idx 28)
        _bg_row(sid, 28, GRY, 14), _bold_row(sid, 28, 14),
        _col_width(sid, 0, 1, 200),
        _col_width(sid, 1, 4, 155),
        _col_width(sid, 4, 14, 140),
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
        # Currency format для B13:D27 (сводка по сотрудникам)
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 12, "endRowIndex": 27,
                      "startColumnIndex": 1, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "CURRENCY", "pattern": "# ##0 \"\u20bd\""},
            }},
            "fields": "userEnteredFormat.numberFormat",
        }},
        # Bold итог (col D = idx 3)
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 12, "endRowIndex": 27,
                      "startColumnIndex": 3, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
            "fields": "userEnteredFormat.textFormat.bold",
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
