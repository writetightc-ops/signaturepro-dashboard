"""
SignaturePro — Дашборд ЗП
Запуск: python dashboard.py
Открыть: http://localhost:5000

Данные: Google Sheets (настроить в config.py)
"""

from flask import Flask, render_template, request, jsonify
from datetime import datetime, date
import os
import re
import time

app = Flask(__name__, template_folder="templates")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# ─── Подключение к Google Sheets ─────────────────────────────────────────────

def _prepare_cloud_credentials():
    """Создаёт credentials.json из переменной окружения GOOGLE_CREDENTIALS_B64."""
    b64 = os.environ.get("GOOGLE_CREDENTIALS_B64", "")
    if not b64:
        return None
    import base64, json
    try:
        data = base64.b64decode(b64).decode("utf-8")
        json.loads(data)
        tmp = os.path.join(BASE_DIR, "_credentials_cloud.json")
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(data)
        return tmp
    except Exception as e:
        print(f"  Ошибка декодирования GOOGLE_CREDENTIALS_B64: {e}")
        return None


try:
    from config import GOOGLE_SHEETS_URL, CREDENTIALS_FILE, CACHE_SECONDS
    _CREDS_PATH = os.path.join(BASE_DIR, CREDENTIALS_FILE)
except ImportError:
    GOOGLE_SHEETS_URL = os.environ.get("GOOGLE_SHEETS_URL", "")
    CREDENTIALS_FILE  = "credentials.json"
    CACHE_SECONDS     = 45
    _CREDS_PATH       = os.path.join(BASE_DIR, CREDENTIALS_FILE)

try:
    from config import MGMT_SHEETS_URL
except ImportError:
    MGMT_SHEETS_URL = os.environ.get("MGMT_SHEETS_URL", "")

if not os.path.exists(_CREDS_PATH):
    cloud_path = _prepare_cloud_credentials()
    if cloud_path:
        _CREDS_PATH = cloud_path

if not GOOGLE_SHEETS_URL:
    raise SystemExit("  Ошибка: задай GOOGLE_SHEETS_URL в config.py")
if not os.path.exists(_CREDS_PATH):
    raise SystemExit(f"  Ошибка: файл учётных данных не найден: {_CREDS_PATH}")

import gsheets as _gs
print(f"  Источник данных: Google Sheets")
print(f"  Таблица заказов: {GOOGLE_SHEETS_URL[:60]}...")
if MGMT_SHEETS_URL:
    print(f"  Таблица руководства: {MGMT_SHEETS_URL[:60]}...")
else:
    print("  Таблица руководства: не настроена (MGMT_SHEETS_URL пустой)")


# ─── Ставки мотивации: встроенный fallback ────────────────────────────────────
#   Используются если таблица руководства недоступна или MGMT_SHEETS_URL не задан.
#   Актуальные ставки хранятся в листе "Ставки" таблицы руководства.

_FALLBACK_CAL_RATES = {
    "E":      {"варианты": 150, "обучение": 200, "бонус_без_правок": 100, "тарифный_бонус":   0, "правка": 75},
    "ST":     {"варианты": 150, "обучение": 200, "бонус_без_правок": 100, "тарифный_бонус":   0, "правка": 75},
    "E_FAST": {"варианты": 150, "обучение": 200, "бонус_без_правок": 150, "тарифный_бонус": 300, "правка": 75},
    "OPTNEW": {"варианты": 500, "обучение": 200, "бонус_без_правок": 200, "тарифный_бонус":   0, "правка": 75},
    "PRNEW":  {"варианты": 500, "обучение": 200, "бонус_без_правок": 250, "тарифный_бонус": 300, "правка": 75},
    "OPT2":   {"варианты": 300, "обучение": 200, "бонус_без_правок": 150, "тарифный_бонус":   0, "правка": 75},
}
_FALLBACK_MGR_RATES = {
    "E": 200, "ST": 200, "E_FAST": 200, "OPTNEW": 300, "PRNEW": 300, "OPT2": 300,
}
_FALLBACK_GRADES = {
    "Марьям":          1.0,
    "Лена Вовина":     1.0,
    "Катерина Попова": 1.0,
    "Катя Дорожкина":  1.0,
    "Ольга Струкова":  1.0,
    "Мария Тимофеева": 1.0,
}

_rates_cache = {"ts": 0, "cal": None, "mgr": None, "grades": None}
_RATES_TTL   = 600  # 10 минут


def _get_rates():
    """
    Загружает ставки из листа 'Ставки' таблицы руководства (кеш 10 мин).
    При ошибке возвращает встроенные значения.
    """
    now = time.time()
    if _rates_cache["cal"] and (now - _rates_cache["ts"]) < _RATES_TTL:
        return _rates_cache["cal"], _rates_cache["mgr"], _rates_cache["grades"]

    if MGMT_SHEETS_URL:
        try:
            cal, mgr, emps = _gs.read_rates(MGMT_SHEETS_URL, _CREDS_PATH, cache_ttl=0)
            if cal:
                grades = {name: d["coefficient"] for name, d in emps.items()}
                _rates_cache.update({"ts": now, "cal": cal, "mgr": mgr, "grades": grades})
                print(f"  Ставки загружены из Google Sheets ({len(cal)} тарифов, {len(emps)} сотрудников)")
                return cal, mgr, grades
        except Exception as e:
            print(f"  Предупреждение: не удалось загрузить ставки из таблицы руководства: {e}")
            print("  Используются встроенные ставки из dashboard.py")

    # Fallback: встроенные ставки
    _rates_cache.update({
        "ts": now,
        "cal": _FALLBACK_CAL_RATES,
        "mgr": _FALLBACK_MGR_RATES,
        "grades": _FALLBACK_GRADES,
    })
    return _FALLBACK_CAL_RATES, _FALLBACK_MGR_RATES, _FALLBACK_GRADES


def _invalidate_rates_cache():
    _rates_cache["ts"] = 0
    _rates_cache["cal"] = None
    if MGMT_SHEETS_URL:
        _gs.invalidate_mgmt_cache(MGMT_SHEETS_URL)


# ─── Утилиты ─────────────────────────────────────────────────────────────────

def is_cyrillic(text: str) -> bool:
    return bool(re.search(r"[а-яА-ЯёЁ]", str(text)))


def fmt_date(d):
    if d is None:
        return ""
    if isinstance(d, datetime):
        d = d.date()
    return d.strftime("%d.%m.%y") if d else ""


# ─── Расчёт заработка по одному заказу ───────────────────────────────────────

def calc_order_earnings(order, date_from, date_to, cal_rates, mgr_rates):
    """
    Считает заработок каллиграфа и менеджера по одному заказу
    за период [date_from, date_to].

    Логика начислений:
    - Варианты + тарифный бонус → когда дата «Варианты» попадает в период.
    - Правки       → каждая правка, чья дата попадает в период.
    - Обучение     → когда дата «Обучение» попадает в период.
    - Бонус без правок → при завершении в периоде И если по заказу
                          ВООБЩЕ нет правок (независимо от периода).
    - Менеджерский бонус → при завершении заказа в периоде.
    """
    tariff = order["тариф"]
    rates = cal_rates.get(tariff)
    if not rates:
        return {}, 0.0, 0.0

    def in_p(d):
        return d is not None and date_from <= d <= date_to

    cal = {}

    if in_p(order["варианты"]):
        cal["варианты"] = rates["варианты"]
        if rates["тарифный_бонус"] > 0:
            cal["тарифный_бонус"] = rates["тарифный_бонус"]

    правки_sum = 0
    for k in ("правка1", "правка2", "правка3", "правка4", "правка5", "правка6"):
        if in_p(order[k]):
            правки_sum += rates["правка"]
    if правки_sum:
        cal["правки"] = правки_sum

    if in_p(order["обучение"]):
        cal["обучение"] = rates["обучение"]

    mgr_bonus = 0.0
    if in_p(order["завершение"]):
        total_правок = sum(
            1 for k in ("правка1", "правка2", "правка3", "правка4", "правка5", "правка6")
            if order[k] is not None
        )
        if total_правок == 0:
            cal["бонус_без_правок"] = rates["бонус_без_правок"]
        mgr_bonus = float(mgr_rates.get(tariff, 0))

    return cal, sum(cal.values()), mgr_bonus


# ─── Подсчёт ЗП всех сотрудников ─────────────────────────────────────────────

def calculate_salary(date_from, date_to):
    """Считает ЗП всех сотрудников за период. Читает ВСЕ листы заказов."""
    try:
        all_orders, order_sheets = _gs.read_all_orders(
            GOOGLE_SHEETS_URL, _CREDS_PATH, CACHE_SECONDS
        )
        cal_rates, mgr_rates, grades = _get_rates()
    except Exception as e:
        return {"error": str(e)}

    results = {}

    def ensure(name, role):
        if name not in results:
            results[name] = {
                "name": name, "role": role,
                "coefficient": grades.get(name, 1.0),
                "total": 0.0,
                "total_usa": 0.0,
                "total_ru": 0.0,
                "breakdown": {},
                "orders_count": 0,
                "orders_done": 0,
                "orders_detail": [],
            }

    for order in all_orders:
        cal_name = order["каллиграф"]
        mgr_name = order["менеджер"]
        if not cal_name and not mgr_name:
            continue

        cal_bd, cal_total, mgr_bonus = calc_order_earnings(
            order, date_from, date_to, cal_rates, mgr_rates
        )
        country = "RU" if is_cyrillic(order["клиент"]) else "USA"

        # ── Каллиграф ───────────────────────────────────────────────────────
        if cal_name and cal_total > 0:
            ensure(cal_name, "Каллиграф")
            coeff   = grades.get(cal_name, 1.0)
            weighted = cal_total * coeff
            results[cal_name]["total"] += weighted
            if country == "RU":
                results[cal_name]["total_ru"] += weighted
            else:
                results[cal_name]["total_usa"] += weighted
            results[cal_name]["orders_count"] += 1
            if order["завершение"] and date_from <= order["завершение"] <= date_to:
                results[cal_name]["orders_done"] += 1
            for k, v in cal_bd.items():
                bd = results[cal_name]["breakdown"]
                bd[k] = bd.get(k, 0) + v * coeff
            results[cal_name]["orders_detail"].append({
                "клиент":      order["клиент"],
                "тариф":       order["тариф"],
                "поступление": fmt_date(order["поступление"]),
                "завершение":  fmt_date(order["завершение"]),
                "варианты":    fmt_date(order["варианты"]),
                "правок":      sum(1 for k in ("правка1","правка2","правка3","правка4","правка5","правка6") if order[k]),
                "обучение":    fmt_date(order["обучение"]),
                "заработок":   round(cal_total * coeff, 2),
                "breakdown":   {k: round(v * coeff, 2) for k, v in cal_bd.items()},
                "лист":        order["_sheet"],
                "страна":      country,
            })

        # ── Менеджер ─────────────────────────────────────────────────────────
        if mgr_name and mgr_bonus > 0:
            ensure(mgr_name, "Менеджер")
            results[mgr_name]["total"] += mgr_bonus
            if country == "RU":
                results[mgr_name]["total_ru"] += mgr_bonus
            else:
                results[mgr_name]["total_usa"] += mgr_bonus
            bd = results[mgr_name]["breakdown"]
            bd["бонус_за_заказ"] = bd.get("бонус_за_заказ", 0) + mgr_bonus
            results[mgr_name]["orders_count"] += 1
            results[mgr_name]["orders_done"]  += 1
            results[mgr_name]["orders_detail"].append({
                "клиент":      order["клиент"],
                "тариф":       order["тариф"],
                "поступление": fmt_date(order["поступление"]),
                "завершение":  fmt_date(order["завершение"]),
                "варианты":    "",
                "правок":      sum(1 for k in ("правка1","правка2","правка3","правка4","правка5","правка6") if order[k]),
                "обучение":    "",
                "заработок":   round(mgr_bonus, 2),
                "breakdown":   {"бонус_за_заказ": round(mgr_bonus, 2)},
                "лист":        order["_sheet"],
                "страна":      country,
            })

    for emp in results.values():
        emp["total"]     = round(emp["total"], 2)
        emp["total_usa"] = round(emp["total_usa"], 2)
        emp["total_ru"]  = round(emp["total_ru"], 2)
        emp["breakdown"] = {k: round(v, 2) for k, v in emp["breakdown"].items() if v > 0}

    return {
        "employees":    results,
        "period_from":  date_from.isoformat(),
        "period_to":    date_to.isoformat(),
        "sheets_used":  order_sheets,
        "total_orders": len(all_orders),
        "rates_source": "gsheets" if MGMT_SHEETS_URL else "builtin",
    }


# ─── Flask маршруты ──────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/meta")
def api_meta():
    try:
        sheets = _gs.get_order_sheet_names(GOOGLE_SHEETS_URL, _CREDS_PATH)
        return jsonify({"sheets": sheets, "source": "gsheets"})
    except Exception as e:
        return jsonify({"sheets": [], "source": "gsheets", "error": str(e)}), 500


@app.route("/api/sheets")
def api_sheets():
    return api_meta()


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    """Сбрасывает кеш Google Sheets и кеш ставок."""
    try:
        _gs.invalidate_cache(GOOGLE_SHEETS_URL)
        _invalidate_rates_cache()
        return jsonify({"ok": True, "message": "Кеш сброшен, данные обновятся из Google Sheets"})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/salary")
def api_salary():
    """GET /api/salary?from=2026-04-01&to=2026-04-30"""
    try:
        today     = date.today()
        from_str  = request.args.get("from")
        to_str    = request.args.get("to")
        date_from = datetime.strptime(from_str, "%Y-%m-%d").date() if from_str else date(today.year, today.month, 1)
        date_to   = datetime.strptime(to_str,   "%Y-%m-%d").date() if to_str   else today

        if date_from > date_to:
            return jsonify({"error": "Дата «с» не может быть позже даты «по»"}), 400

        return jsonify(calculate_salary(date_from, date_to))

    except ValueError as e:
        return jsonify({"error": f"Неверный формат даты: {e}"}), 400
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/api/transfer", methods=["POST"])
def api_transfer():
    """Переносит незавершённые заказы из одного листа в другой."""
    try:
        body       = request.get_json(force=True) or {}
        from_sheet = (body.get("from_sheet") or "").strip()
        to_sheet   = (body.get("to_sheet")   or "").strip()

        if not from_sheet or not to_sheet:
            return jsonify({"error": "Укажи from_sheet и to_sheet"}), 400
        if from_sheet == to_sheet:
            return jsonify({"error": "Исходный и целевой листы совпадают"}), 400

        count = _gs.transfer_orders(GOOGLE_SHEETS_URL, _CREDS_PATH, from_sheet, to_sheet)
        if count == 0:
            return jsonify({"transferred": 0, "message": "Незавершённых заказов нет"})
        return jsonify({
            "transferred": count,
            "from_sheet":  from_sheet,
            "to_sheet":    to_sheet,
            "message":     f"Перенесено {count} заказов из «{from_sheet}» в «{to_sheet}»",
        })

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/api/closed_periods")
def api_closed_periods():
    """Возвращает список зафиксированных периодов из таблицы руководства."""
    try:
        if not MGMT_SHEETS_URL:
            return jsonify({"periods": [], "mgmt_configured": False})
        periods = _gs.read_closed_periods(MGMT_SHEETS_URL, _CREDS_PATH)
        return jsonify({"periods": periods, "mgmt_configured": True})
    except Exception as e:
        return jsonify({"periods": [], "mgmt_configured": True, "error": str(e)})


@app.route("/api/close_period", methods=["POST"])
def api_close_period():
    """
    Фиксирует ЗП за период в таблице руководства (лист История_ЗП).
    POST body: {"from": "2026-05-01", "to": "2026-05-31"}
    """
    try:
        if not MGMT_SHEETS_URL:
            return jsonify({
                "error": "Таблица руководства не настроена. "
                         "Добавь MGMT_SHEETS_URL в config.py."
            }), 400

        body     = request.get_json(force=True) or {}
        from_str = body.get("from", "")
        to_str   = body.get("to",   "")

        if not from_str or not to_str:
            return jsonify({"error": "Укажи from и to в теле запроса"}), 400

        date_from = datetime.strptime(from_str, "%Y-%m-%d").date()
        date_to   = datetime.strptime(to_str,   "%Y-%m-%d").date()
        if date_from > date_to:
            return jsonify({"error": "Дата «с» не может быть позже даты «по»"}), 400

        # Считаем ЗП
        result = calculate_salary(date_from, date_to)
        if "error" in result:
            return jsonify({"error": result["error"]}), 500

        if not result.get("employees"):
            return jsonify({"error": "Нет начислений за выбранный период"}), 400

        # Записываем в История_ЗП
        count = _gs.write_salary_history(
            MGMT_SHEETS_URL, _CREDS_PATH,
            date_from, date_to,
            result["employees"],
        )

        period_str = f"{date_from.strftime('%d.%m.%Y')} \u2014 {date_to.strftime('%d.%m.%Y')}"
        return jsonify({
            "ok":              True,
            "period":          period_str,
            "employees_count": count,
            "message":         f"ЗП зафиксирована: {count} сотрудников · период «{period_str}»",
        })

    except ValueError as e:
        # Период уже закрыт или нет данных
        return jsonify({"error": str(e)}), 409
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


if __name__ == "__main__":
    print("=" * 50)
    print("  SignaturePro — Dashboard ZP")
    print("=" * 50)
    print("\n  Открой в браузере: http://localhost:5000\n")
    app.run(debug=False, host="0.0.0.0", port=5000)
