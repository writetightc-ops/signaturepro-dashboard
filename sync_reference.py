"""
Синхронизация листа «Справочник» и выпадающего списка тарифов (столбец F).

1. Перезаписывает «Справочник» таблицей из 7 тарифов
   (E, ST, E_FAST, OPTNEW 300/400, PRNEW, ДОП_ОБУЧЕНИЕ, ДОП_ПОДПИСЬ),
   таблицей бонусов менеджера и списком сотрудников
   (существующие сотрудники сохраняются).
2. Ставит на столбец F во всех листах заказов + «Шаблон»
   выпадающий список ONE_OF_LIST с этими 7 тарифами
   (заменяет старый список с «OPT²»).

Трогает: «Справочник» и data validation столбца F листов заказов.
НЕ трогает: данные заказов, «Ставки», «История_ЗП», «Корректировки».

Запуск: python sync_reference.py
"""
import os
import sys
import time

import gspread
from google.oauth2.service_account import Credentials

try:
    from config import GOOGLE_SHEETS_URL, CREDENTIALS_FILE
except ImportError:
    GOOGLE_SHEETS_URL = os.environ.get("GOOGLE_SHEETS_URL", "")
    CREDENTIALS_FILE = "credentials.json"

from setup_mgmt_sheet import _rgb, _bold_row, _bg_row, _white_text_row, _col_width, _freeze, _with_retry

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Список листов, которые не являются заказами (их F не трогаем, кроме Шаблона — трогаем)
SKIP_F_SHEETS = {"Справочник", "import", "export", "Дисциплина", "ЗП отчет месяц",
                 "Бонусы", "Рейты", "Сотрудники", "Статистика", "Импорт"}

# 7 тарифов: Варианты, Обучение, Бонус без правок, Тарифный бонус, Итог без правок, Правка
TARIFFS = [
    ["E",            150, 200, 100,   0, 450, 75],
    ["ST",           150, 200, 100,   0, 450, 75],
    ["E_FAST",       150, 200, 150, 300, 800, 75],
    ["OPTNEW",       300, 400, 200,   0, 900, 75],
    ["PRNEW",        500, 200, 250, 300, 1250, 75],
    ["ДОП_ОБУЧЕНИЕ",   0, 200,   0,   0, 200, 0],
    ["ДОП_ПОДПИСЬ",  250, 200, 100,   0, 550, 75],
]
# Бонус менеджера по тарифу
MGR_BONUS = {
    "E": 200, "ST": 200, "E_FAST": 200, "OPTNEW": 300, "PRNEW": 300,
    "ДОП_ОБУЧЕНИЕ": 0, "ДОП_ПОДПИСЬ": 200,
}
TARIFF_VALUES = ["E", "ST", "E_FAST", "OPTNEW", "PRNEW", "ДОП_ОБУЧЕНИЕ", "ДОП_ПОДПИСЬ"]


def _creds_path():
    base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, CREDENTIALS_FILE)


def get_client():
    creds = Credentials.from_service_account_file(_creds_path(), scopes=SCOPES)
    try:
        print(f"  Сервисный аккаунт: {creds.service_account_email}")
    except Exception:
        pass
    return gspread.authorize(creds)


def read_existing_employees(ws):
    """Читает раздел «СОТРУДНИКИ» из Справочника → список строк."""
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
        if not first:
            continue
        r = list(row) + [""] * 5
        emp.append([r[0].strip(), r[1].strip(), r[2].strip(), r[3].strip() or "1", r[4].strip() or "TRUE"])
    return emp


def rewrite_reference(sh, ws, employees):
    print("  Перезаписываю «Справочник»...")
    sid = ws.id

    # Значения
    values = (
        [["ТАРИФЫ И МОТИВАЦИЯ КАЛЛИГРАФОВ"]]
        + [["Тариф", "Варианты", "Обучение", "Бонус без правок",
            "Тарифный бонус", "Итог без правок", "Стоимость 1 правки"]]
        + TARIFFS
        + [[]]
        + [["БОНУС МЕНЕДЖЕРА (за завершённый заказ)"]]
        + [["Тариф", "Бонус менеджера, ₽"]]
        + [[t, MGR_BONUS[t]] for t in TARIFF_VALUES]
        + [[]]
        + [["СОТРУДНИКИ"]]
        + [["Имя", "Роль", "Грейд", "Коэффициент", "Активен"]]
        + employees
    )

    _with_retry(lambda: ws.clear())
    time.sleep(1)
    _with_retry(lambda: ws.update(range_name="A1", values=values, value_input_option="USER_ENTERED"))
    time.sleep(1)

    # Индексы строк (0-based)
    title_r = 0
    tar_hdr = 1
    tar_first = 2
    tar_last = tar_first + len(TARIFFS) - 1          # 8
    mgr_title = tar_last + 2                          # 10
    mgr_hdr = mgr_title + 1                           # 11
    emp_title = mgr_hdr + 1 + len(TARIFF_VALUES) + 1  # 11+1+7+1=20
    emp_hdr = emp_title + 1                          # 21

    BLU_DARK = _rgb(26, 115, 232)
    BLU_LIGHT = _rgb(197, 218, 255)
    GRN_DARK = _rgb(52, 168, 83)
    GRN_LIGHT = _rgb(182, 215, 168)
    YLW = _rgb(251, 188, 4)

    fmt = [
        _bg_row(sid, title_r, BLU_DARK, 7), _white_text_row(sid, title_r, 7),
        _bg_row(sid, tar_hdr, BLU_LIGHT, 7), _bold_row(sid, tar_hdr, 7),
        _bg_row(sid, mgr_title, GRN_DARK, 2), _white_text_row(sid, mgr_title, 2),
        _bg_row(sid, mgr_hdr, GRN_LIGHT, 2), _bold_row(sid, mgr_hdr, 2),
        _bg_row(sid, emp_title, YLW, 5), _bold_row(sid, emp_title, 5),
        _bg_row(sid, emp_hdr, _rgb(241, 241, 241), 5), _bold_row(sid, emp_hdr, 5),
        _freeze(sid, 1),
        _col_width(sid, 0, 1, 150),
        _col_width(sid, 1, 7, 130),
    ]
    _with_retry(lambda: sh.batch_update({"requests": fmt}))
    print("    [OK] Справочник обновлён")


def set_tariff_dropdown(sh, ws):
    """Ставит на F2:F1000 выпадающий список тарифов (заменяет старые правила)."""
    sid = ws.id
    req = {
        "setDataValidation": {
            "range": {
                "sheetId": sid,
                "startRowIndex": 1,
                "endRowIndex": 1000,
                "startColumnIndex": 5,   # F
                "endColumnIndex": 6,
            },
            "rule": {
                "condition": {
                    "type": "ONE_OF_LIST",
                    "values": [{"userEnteredValue": v} for v in TARIFF_VALUES],
                },
                "showCustomUi": True,
                "strict": True,
            },
        }
    }
    _with_retry(lambda: sh.batch_update({"requests": [req]}))
    time.sleep(0.5)


def main():
    print("=" * 60)
    print("  SignaturePro — Синхронизация Справочника и дропдауна тарифов")
    print("=" * 60)
    if not GOOGLE_SHEETS_URL:
        print("  Ошибка: GOOGLE_SHEETS_URL не задан в config.py")
        sys.exit(1)

    gc = get_client()
    sh = _with_retry(lambda: gc.open_by_url(GOOGLE_SHEETS_URL))
    print(f"  Открыта: {sh.title}")
    titles = [w.title for w in sh.worksheets()]
    print(f"  Листы: {titles}")

    # 1. Справочник
    if "Справочник" not in titles:
        print("  Ошибка: лист «Справочник» не найден")
        sys.exit(1)
    ws_ref = sh.worksheet("Справочник")
    employees = read_existing_employees(ws_ref)
    print(f"  Сотрудников в Справочнике: {len(employees)}")
    rewrite_reference(sh, ws_ref, employees)
    time.sleep(1)

    # 2. Дропдаун на столбце F во всех листах заказов + Шаблон
    print("  Ставлю выпадающий список тарифов в столбец F:")
    for w in sh.worksheets():
        if w.title in SKIP_F_SHEETS:
            continue
        print(f"    → {w.title}")
        set_tariff_dropdown(sh, w)
    print("    [OK] Дропдаун обновлён")

    print("\n  7 тарифов в дропдауне:", ", ".join(TARIFF_VALUES))
    print("\n" + "=" * 60)
    print("  [DONE] Справочник и дропдаун обновлены.")
    print("  Данные заказов, Ставки, История_ЗП, Корректировки НЕ затронуты.")
    print("=" * 60)


if __name__ == "__main__":
    main()