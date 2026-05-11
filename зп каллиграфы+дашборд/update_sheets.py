"""
Обновляет листы таблицы руководства (НЕ удаляя данные!):

  1. Корректировки  — выпадающие списки для Периода, Сотрудника, Типа; числовой формат Суммы
  2. Итого_к_выплате — добавляет раздел «СВОДКА ПО СОТРУДНИКАМ»
                       (базовая ЗП + доплаты + итого к выплате за период)

Запуск:  python update_sheets.py
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
    raise RuntimeError(f"Превышен лимит API: {last_err}")


def get_client():
    creds_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), CREDENTIALS_FILE)
    if not os.path.exists(creds_path):
        raise FileNotFoundError(f"credentials.json не найден: {creds_path}")
    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    return gspread.authorize(creds)


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
        "cell": {"userEnteredFormat": {
            "textFormat": {"foregroundColor": _rgb(255, 255, 255), "bold": True}
        }},
        "fields": "userEnteredFormat.textFormat",
    }}


# ═══════════════════════════ 1. КОРРЕКТИРОВКИ ══════════════════════════════════

def update_corrections(sh, ws):
    """
    Добавляет выпадающие списки в лист Корректировки:
      - Период     (A): список уже закрытых периодов из История_ЗП
      - Сотрудник  (B): список сотрудников из Ставки
      - Тип        (C): Премия / Доплата / Вычет / Штраф
      - Сумма      (D): денежный формат
    Удаляет старую строку-подсказку, если она есть.
    """
    print("  Обновляю лист 'Корректировки'...")
    sid = ws.id

    # Удаляем строку-подсказку (строка 2), если она ещё есть
    row2 = ws.row_values(2)
    if row2 and any("Типы:" in str(v) or "Формат периода:" in str(v) for v in row2):
        _with_retry(lambda: ws.delete_rows(2))
        print("    [OK] Удалена строка-подсказка")
        time.sleep(1)

    requests = [
        # A (Период): выпадающий список из История_ЗП!A2:A1000
        {"setDataValidation": {
            "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": 500,
                      "startColumnIndex": 0, "endColumnIndex": 1},
            "rule": {
                "condition": {
                    "type": "ONE_OF_RANGE",
                    "values": [{"userEnteredValue": "='\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!$A$2:$A$1000"}]
                },
                "showCustomUi": True,
                "strict": False,
            },
        }},

        # B (Сотрудник): выпадающий список из Ставки!A12:A100
        {"setDataValidation": {
            "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": 500,
                      "startColumnIndex": 1, "endColumnIndex": 2},
            "rule": {
                "condition": {
                    "type": "ONE_OF_RANGE",
                    "values": [{"userEnteredValue": "='\u0421\u0442\u0430\u0432\u043a\u0438'!$A$12:$A$100"}]
                },
                "showCustomUi": True,
                "strict": False,
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
                "showCustomUi": True,
                "strict": True,
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

    _with_retry(lambda: sh.batch_update({"requests": requests}))
    print("    [OK] Дропдауны: Период, Сотрудник, Тип; формат Суммы")
    time.sleep(1)


# ═══════════════════════════ 2. ИТОГО К ВЫПЛАТЕ ════════════════════════════════

# Формат периода, который совпадает с тем что пишет дашборд в История_ЗП:
# "01.04.2026 — 30.04.2026"

def update_total_sheet(sh, ws):
    """
    Обновляет Итого_к_выплате:
      — Строки 11-12: заголовок + хедеры раздела 'СВОДКА ПО СОТРУДНИКАМ'
      — Строки 13-27: QUERY + SUMIFS (базовая ЗП + доплаты + итого, по каждому сотруднику)
      — Строка  28:   пустая (разделитель)
      — Строки 29-30: заголовки + QUERY таблицы истории/доплат

    Строки 1-10 (сводные суммы) НЕ трогает.
    """
    print("  Обновляю лист 'Итого_к_выплате'...")
    sid = ws.id

    # Очищаем строки 11-35, не трогая 1-10
    _with_retry(lambda: ws.batch_clear(["A11:N35"]))
    time.sleep(1)

    # ── Статические тексты ───────────────────────────────────────────────────
    labels = [
        ("A11", [["\u0421\u0412\u041e\u0414\u041a\u0410 \u041f\u041e \u0421\u041e\u0422\u0420\u0423\u0414\u041d\u0418\u041a\u0410\u041c \u2014 \u0411\u0410\u0417\u041e\u0412\u0410\u042f \u0417\u041f + \u0414\u041e\u041f\u041b\u0410\u0422\u042b"]]),
        ("A12", [["\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a",
                  "\u0411\u0430\u0437\u043e\u0432\u0430\u044f \u0417\u041f",
                  "\u0414\u043e\u043f\u043b\u0430\u0442\u044b (\u00b1)",
                  "\u0418\u0442\u043e\u0433\u043e \u043a \u0432\u044b\u043f\u043b\u0430\u0442\u0435"]]),
        ("A29", [["\u0418\u0421\u0422\u041e\u0420\u0418\u042f \u0417\u041f \u0417\u0410 \u041f\u0415\u0420\u0418\u041e\u0414"]]),
        ("H29", [["\u0414\u041e\u041f\u041b\u0410\u0422\u042b \u0417\u0410 \u041f\u0415\u0420\u0418\u041e\u0414"]]),
    ]

    for cell, val in labels:
        _with_retry(lambda c=cell, v=val: ws.update(
            range_name=c, values=v, value_input_option="USER_ENTERED"
        ))
        time.sleep(0.4)

    # ── Формулы: QUERY для списка сотрудников (A13, спускается вниз) ─────────
    #    Возвращает уникальных сотрудников из История_ЗП для выбранного периода
    a13 = (
        "=IFERROR(QUERY('\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!A:C,"
        "\"SELECT C WHERE A='\"&$C$3&\"' GROUP BY C ORDER BY C\",0),\"\")"
    )

    # Формулы SUMIFS для строк 13..27
    # B = базовая ЗП из История_ЗП (столбец O = "К выплате")
    # C = доплаты из Корректировки (столбец D = "Сумма")
    # D = B + C
    row_data = []
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
        row_data.append((f"B{r}:D{r}", [[b, c, d]]))

    # История ЗП за период (A30) — детализация по сотрудникам
    history_q = (
        "=IFERROR(QUERY('\u0418\u0441\u0442\u043e\u0440\u0438\u044f_\u0417\u041f'!A:Q,"
        "\"SELECT C,D,F,O,P,Q WHERE A='\"&C3&\"' "
        "ORDER BY D,C "
        "LABEL C '\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a',"
        "D '\u0420\u043e\u043b\u044c',F '\u0417\u0430\u0432\u0435\u0440\u0448\u0435\u043d\u043e',"
        "O '\u041a \u0432\u044b\u043f\u043b\u0430\u0442\u0435',P 'USA',Q 'RU'\",0),"
        "\"\u041d\u0435\u0442 \u0434\u0430\u043d\u043d\u044b\u0445\")"
    )

    # Доплаты за период (H30)
    corr_q = (
        "=IFERROR(QUERY('\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438'!A:F,"
        "\"SELECT B,C,D,E WHERE A='\"&C3&\"' "
        "ORDER BY B "
        "LABEL B '\u0421\u043e\u0442\u0440\u0443\u0434\u043d\u0438\u043a',"
        "C '\u0422\u0438\u043f',D '\u0421\u0443\u043c\u043c\u0430',E '\u041a\u043e\u043c\u043c\u0435\u043d\u0442\u0430\u0440\u0438\u0439'\",0),"
        "\"\u041d\u0435\u0442 \u0434\u043e\u043f\u043b\u0430\u0442\")"
    )

    formulas = [
        ("A13", [[a13]]),
        ("A30", [[history_q]]),
        ("H30", [[corr_q]]),
    ] + row_data

    for cell, val in formulas:
        _with_retry(lambda c=cell, v=val: ws.update(
            range_name=c, values=v, value_input_option="USER_ENTERED"
        ))
        time.sleep(0.4)

    # ── Форматирование ───────────────────────────────────────────────────────
    GRN_DARK  = _rgb(52, 168, 83)
    GRN_LIGHT = _rgb(182, 215, 168)
    GRY       = _rgb(241, 241, 241)

    CURRENCY_FMT = {
        "numberFormat": {"type": "CURRENCY", "pattern": "# ##0 \"\u20bd\""},
    }
    CURRENCY_BOLD = {
        "numberFormat": {"type": "CURRENCY", "pattern": "# ##0 \"\u20bd\""},
        "textFormat": {"bold": True},
    }

    fmt = [
        # Заголовок раздела (строка 11, idx 10)
        _bg_row(sid, 10, GRN_DARK, 5),
        _white_text_row(sid, 10, 5),

        # Хедер колонок (строка 12, idx 11)
        _bg_row(sid, 11, GRN_LIGHT, 5),
        _bold_row(sid, 11, 5),

        # Заголовки истории/доплат (строка 29, idx 28)
        _bg_row(sid, 28, GRY, 14),
        _bold_row(sid, 28, 14),

        # Денежный формат: B13:D27 (столбцы 1-3, строки 12-26)
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 12, "endRowIndex": 27,
                      "startColumnIndex": 1, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": CURRENCY_FMT},
            "fields": "userEnteredFormat.numberFormat",
        }},

        # Жирный итог (D столбец, строки 13-27)
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 12, "endRowIndex": 27,
                      "startColumnIndex": 3, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": CURRENCY_BOLD},
            "fields": "userEnteredFormat(numberFormat,textFormat)",
        }},

        # Ширина столбцов
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 200},
            "fields": "pixelSize",
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": 1, "endIndex": 4},
            "properties": {"pixelSize": 150},
            "fields": "pixelSize",
        }},
    ]
    _with_retry(lambda: sh.batch_update({"requests": fmt}))
    print("    [OK] Раздел 'СВОДКА ПО СОТРУДНИКАМ' добавлен (строки 11-30)")
    time.sleep(1)


# ═══════════════════════════ MAIN ═════════════════════════════════════════════

def main():
    print("=" * 60)
    print("  SignaturePro — Обновление листов таблицы руководства")
    print("=" * 60)

    print("\nПодключение...")
    gc = get_client()
    sh = _with_retry(lambda: gc.open_by_url(MGMT_URL))
    print(f"  Открыта: {sh.title}")

    ws_adj   = sh.worksheet("\u041a\u043e\u0440\u0440\u0435\u043a\u0442\u0438\u0440\u043e\u0432\u043a\u0438")
    ws_total = sh.worksheet("\u0418\u0442\u043e\u0433\u043e_\u043a_\u0432\u044b\u043f\u043b\u0430\u0442\u0435")

    update_corrections(sh, ws_adj)
    update_total_sheet(sh, ws_total)

    print("\n" + "=" * 60)
    print("  [DONE] Листы обновлены!")
    print("  Обнови страницу таблицы в браузере.")
    print("=" * 60)


if __name__ == "__main__":
    main()
