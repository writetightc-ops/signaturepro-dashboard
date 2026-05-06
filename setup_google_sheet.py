"""
Скрипт применения правок к Google Sheets:
  1. Условное форматирование по тарифам (зелёный > жёлтый > серый)
  2. Data validation (тип «Дата») для столбцов с датами
  3. Заголовки новых столбцов T–W (ОС 5, Правка 5, ОС 6, Правка 6)
  4. Формула «Бонус» в столбце X

Запуск:
  python setup_google_sheet.py

Требует credentials.json рядом со скриптом и пакеты:
  pip install gspread google-auth
"""

import time
import os

import gspread
from google.oauth2.service_account import Credentials

# ─── Настройки ────────────────────────────────────────────────────────────────

try:
    from config import GOOGLE_SHEETS_URL, CREDENTIALS_FILE
except ImportError:
    GOOGLE_SHEETS_URL = os.environ.get(
        "GOOGLE_SHEETS_URL",
        "https://docs.google.com/spreadsheets/d/1zVxYAVIXR4cwuknI8lS8wElmRlB2cCJD3ceOogMkWOg/edit",
    )
    CREDENTIALS_FILE = "credentials.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SKIP_SHEETS = {
    "Справочник", "Шаблон", "import", "export", "Дисциплина",
    "ЗП отчет месяц", "Бонусы", "Рейты", "Сотрудники",
    "Статистика", "Импорт", "Sheet1", "Лист1", "Sheet",
}

# 0-based индексы столбцов с датами: A(0), H(7), I(8), J(9), K(10),
#   L(11), M(12), N(13), O(14), P(15), R(17), T(19), U(20), V(21), W(22)
DATE_COL_INDICES = [0, 7, 8, 9, 10, 11, 12, 13, 14, 15, 17, 19, 20, 21, 22]

MAX_DATA_ROWS = 500   # максимум строк данных на листе

NEW_COL_HEADERS = ["ОС 5", "Правка 5", "ОС 6", "Правка 6", "Бонус", "Итого ЗП"]  # T–Y


# ─── Подключение ──────────────────────────────────────────────────────────────

def _with_retry(fn, max_retries=6):
    """Повторяет вызов при ошибке лимита API (429 / RESOURCE_EXHAUSTED)."""
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
    raise RuntimeError(
        f"Превышен лимит запросов Google Sheets API после {max_retries} попыток: {last_err}"
    )


def get_client():
    creds_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), CREDENTIALS_FILE)
    if not os.path.exists(creds_path):
        raise FileNotFoundError(
            f"Файл учётных данных не найден: {creds_path}\n"
            "Скачай из Google Cloud Console → Service Account → Keys → JSON"
        )
    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    return gspread.authorize(creds)


# ─── 1. Условное форматирование ───────────────────────────────────────────────

def _rgb(r, g, b):
    return {"red": r / 255, "green": g / 255, "blue": b / 255}


def apply_conditional_formatting(sh, ws):
    """
    Добавляет три правила условного форматирования.
    Порядок (index) определяет приоритет — меньший индекс = выше приоритет.
    """
    sid = ws.id
    grid = {
        "sheetId": sid,
        "startRowIndex": 1,             # строка 2 (пропускаем заголовок)
        "endRowIndex": MAX_DATA_ROWS + 1,
        "startColumnIndex": 0,          # A
        "endColumnIndex": 24,           # A–X
    }

    requests = [
        # Приоритет 0 — Статус = TRUE → зелёный (#C6EFCE)
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [grid],
                    "booleanRule": {
                        "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [{"userEnteredValue": "=$B2=TRUE"}],
                        },
                        "format": {"backgroundColor": _rgb(198, 239, 206)},
                    },
                },
                "index": 0,
            }
        },
        # Приоритет 1 — Тариф E / ST / E_FAST → светло-жёлтый (#FFFFCC)
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [grid],
                    "booleanRule": {
                        "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [{"userEnteredValue": '=OR($F2="E",$F2="ST",$F2="E_FAST")'}],
                        },
                        "format": {"backgroundColor": _rgb(255, 255, 204)},
                    },
                },
                "index": 1,
            }
        },
        # Приоритет 2 — Тариф OPTNEW / OPT2 / PRNEW → светло-серый (#F0F0F0)
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [grid],
                    "booleanRule": {
                        "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [{"userEnteredValue": '=OR($F2="OPTNEW",$F2="OPT2",$F2="PRNEW")'}],
                        },
                        "format": {"backgroundColor": _rgb(240, 240, 240)},
                    },
                },
                "index": 2,
            }
        },
    ]

    _with_retry(lambda: sh.batch_update({"requests": requests}))
    print("    [OK] Условное форматирование добавлено")
    time.sleep(2)


# ─── 2. Data validation (дата) ────────────────────────────────────────────────

def apply_date_validation(sh, ws):
    """Устанавливает проверку данных «дата» для всех столбцов с датами."""
    sid = ws.id
    requests = []
    for col_idx in DATE_COL_INDICES:
        requests.append({
            "setDataValidation": {
                "range": {
                    "sheetId": sid,
                    "startRowIndex": 1,
                    "endRowIndex": MAX_DATA_ROWS + 1,
                    "startColumnIndex": col_idx,
                    "endColumnIndex": col_idx + 1,
                },
                "rule": {
                    "condition": {"type": "DATE_IS_VALID"},
                    "inputMessage": "Выберите дату",
                    "strict": False,
                    "showCustomUi": True,
                },
            }
        })

    # Разбиваем на батчи по 10, чтобы не переполнить квоту
    for i in range(0, len(requests), 10):
        batch = requests[i:i + 10]
        _with_retry(lambda b=batch: sh.batch_update({"requests": b}))
        time.sleep(1)

    print("    [OK] Data validation (дата) применена")
    time.sleep(1)


# ─── 3 & 4. Новые столбцы T–X ────────────────────────────────────────────────

def _col_letter(idx_0based):
    """Число 0-based → буква столбца (0=A, 25=Z, 26=AA, ...)."""
    result = ""
    n = idx_0based
    while True:
        result = chr(ord("A") + n % 26) + result
        n = n // 26 - 1
        if n < 0:
            break
    return result


def add_new_column_headers(ws):
    """Прописывает заголовки T1:Y1 если их ещё нет."""
    headers = ws.row_values(1)
    while len(headers) < 25:
        headers.append("")

    needs_update = any(headers[19 + i] != name for i, name in enumerate(NEW_COL_HEADERS))

    if needs_update:
        _with_retry(lambda: ws.update(
            range_name="T1:Y1",
            values=[NEW_COL_HEADERS],
            value_input_option="USER_ENTERED",
        ))
        print("    [OK] Заголовки T-Y обновлены")
    else:
        print("    [-] Заголовки T-Y уже установлены")
    time.sleep(1)


def _last_data_row(ws):
    """Возвращает номер последней строки с данными (>= 2)."""
    all_values = _with_retry(ws.get_all_values)
    data_rows = [
        r for r in all_values[1:]
        if (len(r) > 0 and r[0].strip()) or (len(r) > 2 and r[2].strip())
    ]
    return 1 + len(data_rows) if data_rows else 0


def add_bonus_formula(ws):
    """
    Столбец X «Бонус»: бонус без правок — выплачивается ТОЛЬКО если
    заказ завершён (B=TRUE) И в строке нет ни одной правки (I,K,M,O,U,W пусты).
    Значения по тарифу: E/ST=100, E_FAST=150, OPTNEW=200, PRNEW=250.
    """
    ldr = _last_data_row(ws)
    if ldr < 2:
        print("    [-] Строк с данными нет - формула бонуса не добавляется")
        return

    formulas = []
    for row_num in range(2, ldr + 1):
        # Правки 1–6 находятся в столбцах I(9), K(11), M(13), O(15), U(21), W(23)
        has_no_правки = (
            f'COUNTA(I{row_num},K{row_num},M{row_num},'
            f'O{row_num},U{row_num},W{row_num})=0'
        )
        f = (
            f'=IF(AND(B{row_num}=TRUE,{has_no_правки}),'
            f'IF(OR(F{row_num}="E",F{row_num}="ST"),100,'
            f'IF(F{row_num}="E_FAST",150,'
            f'IF(F{row_num}="OPTNEW",200,'
            f'IF(F{row_num}="PRNEW",250,0)))),0)'
        )
        formulas.append([f])

    end_x = ldr  # захватываем в замыкание
    _with_retry(lambda: ws.update(
        range_name=f"X2:X{end_x}",
        values=formulas,
        value_input_option="USER_ENTERED",
    ))
    print(f"    [OK] Формула бонуса (бонус без правок) добавлена в X2:X{ldr}")
    time.sleep(1)


def add_total_formula(ws):
    """
    Столбец Y «Итого ЗП»: полная сумма заработка каллиграфа по строке.
    Включает: варианты + правки×75 + обучение + тарифный_бонус + бонус_без_правок.

    Тарифные ставки:
      Варианты : E/ST/E_FAST=150, OPTNEW/PRNEW=500
      Правки   : 75 за каждую (I,K,M,O,U,W)
      Обучение : 200 (если P заполнена)
      Тарифный бонус: E_FAST/PRNEW=300 при завершении
      Бонус без правок: формула из X
    """
    ldr = _last_data_row(ws)
    if ldr < 2:
        print("    [-] Строк с данными нет - формула Итого не добавляется")
        return

    formulas = []
    for row_num in range(2, ldr + 1):
        r = row_num
        # Варианты
        var = (
            f'IF(G{r}<>"",'
            f'IF(OR(F{r}="E",F{r}="ST",F{r}="E_FAST"),150,'
            f'IF(OR(F{r}="OPTNEW",F{r}="PRNEW"),500,0)),0)'
        )
        # Правки 1–6
        правки = f'COUNTA(I{r},K{r},M{r},O{r},U{r},W{r})*75'
        # Обучение
        обучение = f'IF(P{r}<>"",200,0)'
        # Тарифный бонус (только при завершении)
        тар_бонус = (
            f'IF(AND(B{r}=TRUE,OR(F{r}="E_FAST",F{r}="PRNEW")),300,0)'
        )
        # Бонус без правок — ссылаемся на уже рассчитанный X
        бонус_бп = f'X{r}'

        f = f'={var}+{правки}+{обучение}+{тар_бонус}+{бонус_бп}'
        formulas.append([f])

    end_y = ldr  # захватываем в замыкание
    _with_retry(lambda: ws.update(
        range_name=f"Y2:Y{end_y}",
        values=formulas,
        value_input_option="USER_ENTERED",
    ))
    print(f"    [OK] Формула \"Итого ЗП\" добавлена в Y2:Y{ldr}")
    time.sleep(1)


# ─── Основной цикл ────────────────────────────────────────────────────────────

def main():
    print("Подключение к Google Sheets...")
    gc = get_client()
    sh = _with_retry(lambda: gc.open_by_url(GOOGLE_SHEETS_URL))
    print(f"Открыта таблица: {sh.title}\n")

    month_sheets = [ws for ws in sh.worksheets() if ws.title not in SKIP_SHEETS]

    if not month_sheets:
        print("Не найдено листов с заказами (все листы в SKIP_SHEETS).")
        return

    print(f"Найдено листов для обработки: {len(month_sheets)}")
    for ws in month_sheets:
        print(f"\n>> Лист: {ws.title}")

        # 1. Условное форматирование
        apply_conditional_formatting(sh, ws)

        # 2. Data validation для дат
        apply_date_validation(sh, ws)

        # 3. Заголовки T–X
        add_new_column_headers(ws)

        # 4. Формула бонуса в X
        add_bonus_formula(ws)

        # 5. Формула «Итого ЗП» в Y
        add_total_formula(ws)

    print("\n[DONE] Все правки применены к Google Sheets!")
    print("   Обнови страницу таблицы в браузере чтобы увидеть изменения.")


if __name__ == "__main__":
    main()
