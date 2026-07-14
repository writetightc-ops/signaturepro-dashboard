"""
Скрипт применения правок к Google Sheets:
  1. Условное форматирование по тарифам (зелёный > жёлтый > серый)
  2. Снятие data validation с дат (убирает красные уголки)
  3. Заголовки T-X: ОС 5, Правка 5, ОС 6, Правка 6, Итого ЗП
  4. Формула "Итого ЗП" в столбце X (варианты + правки + обучение + бонусы)
  5. Очистка старого столбца Y (если был)

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

# 0-based индексы столбцов с датами (A=0, G=6 варианты, H=7 ОС1, ...)
DATE_COL_INDICES = [0, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 17, 19, 20, 21, 22]

MAX_DATA_ROWS = 500

# T(19)=ОС 5, U(20)=Правка 5, V(21)=ОС 6, W(22)=Правка 6, X(23)=Итого ЗП
NEW_COL_HEADERS = ["ОС 5", "Правка 5", "ОС 6", "Правка 6", "Итого ЗП"]  # T-X


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
        raise FileNotFoundError(f"Файл учётных данных не найден: {creds_path}")
    creds = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    return gspread.authorize(creds)


# ─── 1. Условное форматирование ───────────────────────────────────────────────

def _rgb(r, g, b):
    return {"red": r / 255, "green": g / 255, "blue": b / 255}


def apply_conditional_formatting(sh, ws):
    sid = ws.id
    grid = {
        "sheetId": sid,
        "startRowIndex": 1,
        "endRowIndex": MAX_DATA_ROWS + 1,
        "startColumnIndex": 0,
        "endColumnIndex": 24,  # A-X
    }

    requests = [
        # Приоритет 0: Статус TRUE -> зелёный
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [grid],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA",
                                      "values": [{"userEnteredValue": "=$B2=TRUE"}]},
                        "format": {"backgroundColor": _rgb(198, 239, 206)},
                    },
                },
                "index": 0,
            }
        },
        # Приоритет 1: E / ST / E_FAST -> светло-жёлтый
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [grid],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA",
                                      "values": [{"userEnteredValue": '=OR($F2="E",$F2="ST",$F2="E_FAST")'}]},
                        "format": {"backgroundColor": _rgb(255, 255, 204)},
                    },
                },
                "index": 1,
            }
        },
        # Приоритет 2: OPTNEW / PRNEW / ДОП_ОБУЧЕНИЕ / ДОП_ПОДПИСЬ -> светло-серый
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [grid],
                    "booleanRule": {
                        "condition": {"type": "CUSTOM_FORMULA",
                                      "values": [{"userEnteredValue": '=OR($F2="OPTNEW",$F2="PRNEW",$F2="ДОП_ОБУЧЕНИЕ",$F2="ДОП_ПОДПИСЬ")'}]},
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


# ─── 2. Data validation на датах: выпадающий календарь, без жёсткой блокировки ─

def apply_date_validation(sh, ws):
    """
    Устанавливает data validation DATE_IS_VALID (календарь) + формат dd.mm
    на все столбцы с датами.
    """
    sid = ws.id
    requests = []
    for col_idx in DATE_COL_INDICES:
        date_range = {
            "sheetId": sid,
            "startRowIndex": 1,
            "endRowIndex": MAX_DATA_ROWS + 1,
            "startColumnIndex": col_idx,
            "endColumnIndex": col_idx + 1,
        }
        requests.append({
            "setDataValidation": {
                "range": date_range,
                "rule": {
                    "condition": {"type": "DATE_IS_VALID"},
                    "showCustomUi": True,
                    "strict": False,
                },
            }
        })
        requests.append({
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

    for i in range(0, len(requests), 10):
        batch = requests[i:i + 10]
        _with_retry(lambda b=batch: sh.batch_update({"requests": b}))
        time.sleep(1)

    print("    [OK] Data validation (календарь) + формат dd.mm применены ко всем датам")
    time.sleep(1)


# ─── 3. Заголовки T-X ────────────────────────────────────────────────────────

def add_new_column_headers(ws):
    """Прописывает заголовки T1:X1."""
    headers = ws.row_values(1)
    while len(headers) < 24:
        headers.append("")

    needs_update = any(headers[19 + i] != name for i, name in enumerate(NEW_COL_HEADERS))

    if needs_update:
        _with_retry(lambda: ws.update(
            range_name="T1:X1",
            values=[NEW_COL_HEADERS],
            value_input_option="USER_ENTERED",
        ))
        print("    [OK] Заголовки T-X обновлены")
    else:
        print("    [-] Заголовки T-X уже установлены")
    time.sleep(1)


# ─── 4. Формула "Итого ЗП" в столбце X ───────────────────────────────────────

def _last_data_row(ws):
    all_values = _with_retry(ws.get_all_values)
    data_rows = [
        r for r in all_values[1:]
        if (len(r) > 0 and r[0].strip()) or (len(r) > 2 and r[2].strip())
    ]
    return 1 + len(data_rows) if data_rows else 0


def add_total_formula(ws):
    """
    Столбец X "Итого ЗП" — полная сумма заработка каллиграфа по строке:
      варианты + правки + обучение + тарифный_бонус + бонус_без_правок

    Логика:
      - Варианты: по тарифу, когда заполнена дата «Варианты» (G).
      - Обучение: по тарифу, когда заполнена дата «Обучение» (P).
      - Правки: COUNTA правок × ставку тарифа.
      - Тарифный бонус: вместе с вариантами (G) для E_FAST / PRNEW.
      - Бонус без правок: при B=TRUE, G заполнено и 0 правок — по тарифу.
    Ставки захардкожены — синхронизируются вручную с листом «Ставки»,
    _FALLBACK_CAL_RATES в dashboard.py и _build_total_formula в gsheets.py.
    """
    ldr = _last_data_row(ws)
    if ldr < 2:
        print("    [-] Строк с данными нет - формула Итого не добавляется")
        return

    formulas = []
    for r in range(2, ldr + 1):
        # Варианты: (тариф → ставка)
        var = (
            f'IF(G{r}<>"",'
            f'IF(OR(F{r}="E",F{r}="ST",F{r}="E_FAST"),150,'
            f'IF(F{r}="OPTNEW",300,'
            f'IF(F{r}="PRNEW",500,'
            f'IF(F{r}="ДОП_ПОДПИСЬ",250,'
            f'IF(F{r}="ДОП_ОБУЧЕНИЕ",0,0))))),0)'
        )
        # Правки: ставка 75 для всех, кроме ДОП_ОБУЧЕНИЕ (0)
        pravki = (
            f'COUNTA(I{r},K{r},M{r},O{r},U{r},W{r})*'
            f'IF(F{r}="ДОП_ОБУЧЕНИЕ",0,75)'
        )
        # Обучение: (тариф → ставка)
        obuchen = (
            f'IF(P{r}<>"",'
            f'IF(F{r}="OPTNEW",400,'
            f'IF(F{r}="ДОП_ОБУЧЕНИЕ",200,200)),0)'
        )
        # Тарифный бонус — вместе с вариантами (G)
        tar_bonus = f'IF(AND(G{r}<>"",OR(F{r}="E_FAST",F{r}="PRNEW")),300,0)'
        no_pravki_check = f'COUNTA(I{r},K{r},M{r},O{r},U{r},W{r})=0'
        # Бонус без правок — при B=TRUE, G заполнено, 0 правок
        bonus_bp = (
            f'IF(AND(B{r}=TRUE,G{r}<>"",{no_pravki_check}),'
            f'IF(OR(F{r}="E",F{r}="ST"),100,'
            f'IF(F{r}="E_FAST",150,'
            f'IF(F{r}="OPTNEW",200,'
            f'IF(F{r}="PRNEW",250,'
            f'IF(F{r}="ДОП_ПОДПИСЬ",100,'
            f'IF(F{r}="ДОП_ОБУЧЕНИЕ",0,0)))))),0)'
        )
        formula = f'={var}+{pravki}+{obuchen}+{tar_bonus}+{bonus_bp}'
        formulas.append([formula])

    end_x = ldr
    _with_retry(lambda: ws.update(
        range_name=f"X2:X{end_x}",
        values=formulas,
        value_input_option="USER_ENTERED",
    ))
    print(f"    [OK] Формула \"Итого ЗП\" добавлена в X2:X{ldr}")
    time.sleep(1)


# ─── 5. Очистка старого столбца Y ────────────────────────────────────────────

def clear_old_y_column(ws):
    """Очищает столбец Y (там был старый «Итого ЗП» или «Бонус»)."""
    ldr = _last_data_row(ws)
    if ldr < 1:
        return

    empty_header = [[""]]
    _with_retry(lambda: ws.update(
        range_name="Y1:Y1",
        values=empty_header,
        value_input_option="USER_ENTERED",
    ))

    if ldr >= 2:
        empty_data = [[""] for _ in range(2, ldr + 1)]
        end_y = ldr
        _with_retry(lambda: ws.update(
            range_name=f"Y2:Y{end_y}",
            values=empty_data,
            value_input_option="USER_ENTERED",
        ))

    print("    [OK] Старый столбец Y очищен")
    time.sleep(1)


# ─── Основной цикл ────────────────────────────────────────────────────────────

def main():
    print("Подключение к Google Sheets...")
    gc = get_client()
    sh = _with_retry(lambda: gc.open_by_url(GOOGLE_SHEETS_URL))
    print(f"Открыта таблица: {sh.title}\n")

    month_sheets = [ws for ws in sh.worksheets() if ws.title not in SKIP_SHEETS]

    if not month_sheets:
        print("Не найдено листов с заказами.")
        return

    print(f"Найдено листов для обработки: {len(month_sheets)}")
    for ws in month_sheets:
        print(f"\n>> Лист: {ws.title}")

        # 1. Условное форматирование
        apply_conditional_formatting(sh, ws)

        # 2. Data validation с календарём (strict=False — без жёсткой блокировки)
        apply_date_validation(sh, ws)

        # 3. Заголовки T-X
        add_new_column_headers(ws)

        # 4. Формула "Итого ЗП" в X (без отдельного "Бонус")
        add_total_formula(ws)

        # 5. Очистить старый Y
        clear_old_y_column(ws)

    print("\n[DONE] Все правки применены к Google Sheets!")
    print("   Обнови страницу таблицы в браузере чтобы увидеть изменения.")


if __name__ == "__main__":
    main()
