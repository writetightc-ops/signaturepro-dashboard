"""
Офлайн-проверка новой логики начислений.
Не требует creds/сети: импортирует только calc_order_earnings и fallback-ставки.
Запуск: python test_logic_offline.py
"""
import os, json, base64, sys
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# Подменяем окружение ДО импорта dashboard (иначе SystemExit на старте).
os.environ["GOOGLE_SHEETS_URL"] = "https://docs.google.com/spreadsheets/d/dummy/edit"
os.environ["MGMT_SHEETS_URL"] = ""  # чтобы ставки брались из fallback
os.environ["GOOGLE_CREDENTIALS_B64"] = base64.b64encode(
    json.dumps({"type": "service_account"}).encode()
).decode()

from datetime import date
from dashboard import calc_order_earnings, _FALLBACK_CAL_RATES, _FALLBACK_MGR_RATES

CAL, MGR = _FALLBACK_CAL_RATES, _FALLBACK_MGR_RATES
APR = (date(2026, 4, 1), date(2026, 4, 30))
MAY = (date(2026, 5, 1), date(2026, 5, 31))

def base_order(**kw):
    o = {
        "поступление": None, "статус": False, "менеджер": "", "каллиграф": "",
        "клиент": "", "тариф": "", "варианты": None,
        "ос1": None, "правка1": None, "ос2": None, "правка2": None,
        "ос3": None, "правка3": None, "ос4": None, "правка4": None,
        "обучение": None, "ссылка": "", "завершение": None, "заметки": "",
        "ос5": None, "правка5": None, "ос6": None, "правка6": None,
    }
    o.update(kw)
    return o

passed = 0
failed = 0

def check(name, cond, extra=""):
    global passed, failed
    if cond:
        passed += 1
        print(f"  [OK] {name}")
    else:
        failed += 1
        print(f"  [FAIL] {name} {extra}")

print("=" * 60)
print("  Офлайн-тест логики начислений")
print("=" * 60)

# 1. OPTNEW: варианты 300, обучение 400
o = base_order(тариф="OPTNEW", каллиграф="Марьям", менеджер="Ольга",
               клиент="Иван Иванов", варианты=date(2026, 4, 10),
               обучение=date(2026, 4, 20), статус=True, завершение=date(2026, 4, 25))
cal, total, mgr = calc_order_earnings(o, *APR, CAL, MGR)
check("OPTNEW варианты=300", cal.get("варианты") == 300, f"got {cal.get('варианты')}")
check("OPTNEW обучение=400", cal.get("обучение") == 400, f"got {cal.get('обучение')}")
check("OPTNEW тарифный_бонус отсутствует", "тарифный_бонус" not in cal)
check("OPTNEW бонус_без_правок=200 (B+R, 0 правок)",
      cal.get("бонус_без_правок") == 200, f"got {cal.get('бонус_без_правок')}")
check("OPTNEW менеджер=300", mgr == 300, f"got {mgr}")
check("OPTNEW total = 300+400+200 = 900", total == 900, f"got {total}")

# 2. OPTNEW обучение НЕ в периоде → только варианты + завершающие
o = base_order(тариф="OPTNEW", каллиграф="Марьям", менеджер="Ольга",
               клиент="А", варианты=date(2026, 4, 10), обучение=date(2026, 5, 15),
               статус=True, завершение=date(2026, 4, 25))
cal, total, mgr = calc_order_earnings(o, *APR, CAL, MGR)
check("OPTNEW обучение в мае → в апреле обучения нет", "обучение" not in cal)
check("OPTNEW менеджер по завершению в периоде (R в апреле)", mgr == 300)

# 3. B=FALSE → нет бонус_без_правок и нет менеджерского, даже если R заполнен
o = base_order(тариф="E", каллиграф="Марьям", менеджер="Ольга", клиент="Б",
               варианты=date(2026, 4, 5), статус=False, завершение=date(2026, 4, 25))
cal, total, mgr = calc_order_earnings(o, *APR, CAL, MGR)
check("B=FALSE → бонус_без_правок нет", "бонус_без_правок" not in cal)
check("B=FALSE → менеджер бонус = 0", mgr == 0)

# 4. E_FAST: тарифный_бонус=300 вместе с вариантами
o = base_order(тариф="E_FAST", каллиграф="Лена", менеджер="Ольга", клиент="В",
               варианты=date(2026, 4, 5), статус=True, завершение=date(2026, 4, 20))
cal, total, mgr = calc_order_earnings(o, *APR, CAL, MGR)
check("E_FAST тарифный_бонус=300 с вариантами", cal.get("тарифный_бонус") == 300)
check("E_FAST бонус_без_правок=150", cal.get("бонус_без_правок") == 150)

# 5. Правка в периоде → бонус_без_правок гасится
o = base_order(тариф="E", каллиграф="Лена", менеджер="Ольга", клиент="Г",
               варианты=date(2026, 4, 5), правка1=date(2026, 4, 12),
               статус=True, завершение=date(2026, 4, 25))
cal, total, mgr = calc_order_earnings(o, *APR, CAL, MGR)
check("Правка в периоде → правки=75", cal.get("правки") == 75)
check("Есть правка → бонус_без_правок нет", "бонус_без_правок" not in cal)
check("Правка не гасит менеджерский бонус", mgr == 200)

# 6. ДОП_ОБУЧЕНИЕ: только обучение=200, менеджер=0 всегда
o = base_order(тариф="ДОП_ОБУЧЕНИЕ", каллиграф="Лена", менеджер="Ольга", клиент="Д",
               обучение=date(2026, 4, 15), статус=True, завершение=date(2026, 4, 20))
cal, total, mgr = calc_order_earnings(o, *APR, CAL, MGR)
check("ДОП_ОБУЧЕНИЕ обучение=200", cal.get("обучение") == 200)
check("ДОП_ОБУЧЕНИЕ вариантов нет (ставка 0)", cal.get("варианты", 0) == 0)
check("ДОП_ОБУЧЕНИЕ бонус_без_правок нет", "бонус_без_правок" not in cal)
check("ДОП_ОБУЧЕНИЕ менеджер = 0", mgr == 0)

# 7. ДОП_ПОДПИСЬ: варианты 250, обучение 200, бп 100, менеджер 200
o = base_order(тариф="ДОП_ПОДПИСЬ", каллиграф="Лена", менеджер="Ольга", клиент="Е",
               варианты=date(2026, 4, 5), обучение=date(2026, 4, 18),
               статус=True, завершение=date(2026, 4, 25))
cal, total, mgr = calc_order_earnings(o, *APR, CAL, MGR)
check("ДОП_ПОДПИСЬ варианты=250", cal.get("варианты") == 250)
check("ДОП_ПОДПИСЬ обучение=200", cal.get("обучение") == 200)
check("ДОП_ПОДПИСЬ бонус_без_правок=100", cal.get("бонус_без_правок") == 100)
check("ДОП_ПОДПИСЬ менеджер=200", mgr == 200)

# 8. Передача другому каллиграфу: варианты в апреле (1-й), завершение в мае (2-й)
o1 = base_order(тариф="OPTNEW", каллиграф="Марьям", менеджер="Ольга", клиент="Клиент Х",
                варианты=date(2026, 4, 10))  # 1-й каллиграф, только варианты
cal1, t1, m1 = calc_order_earnings(o1, *APR, CAL, MGR)
check("1-й каллиграф: варианты в апреле=300", cal1.get("варианты") == 300)
check("1-й каллиграф: в апреле менеджер=0 (нет завершения)", m1 == 0)
o2 = base_order(тариф="OPTNEW", каллиграф="Лена", менеджер="Ольга", клиент="Клиент Х",
                обучение=date(2026, 5, 5), статус=True, завершение=date(2026, 5, 20))
cal2, t2, m2 = calc_order_earnings(o2, *MAY, CAL, MGR)
check("2-й каллиграф: обучение в мае=400", cal2.get("обучение") == 400)
check("2-й каллиграф: менеджер по завершению в мае=300", m2 == 300)

# 9. Неизвестный тариф → пусто
o = base_order(тариф="NOPE", каллиграф="Лена", менеджер="Ольга", клиент="Ж",
               варианты=date(2026, 4, 5))
cal, total, mgr = calc_order_earnings(o, *APR, CAL, MGR)
check("Неизвестный тариф → пустой cal", cal == {})
check("Неизвестный тариф → total=0, mgr=0", total == 0 and mgr == 0)

print()
print("-" * 60)
print("  Тест дедупликации менеджерского бонуса (calculate_salary)")
print("-" * 60)

# Подменяем источник данных и ставки — сети нет.
import dashboard
GRADES = {"Марьям": 1.0, "Лена": 1.0, "Ольга": 1.0}
dashboard._get_rates = lambda: (CAL, MGR, GRADES)

def dup_order(client, variant_date=None, complete=None, mgr="Ольга", cal="Марьям",
              link="", status=True, tarif="OPTNEW"):
    return base_order(
        тариф=tarif, каллиграф=cal, менеджер=mgr, клиент=client,
        поступление=date(2026, 4, 1), варианты=variant_date,
        обучение=complete, завершение=complete, статус=status,
        ссылка=link, _sheet="Апрель_2026",
    )

# Сценарий: клиент передан другому каллиграфу — две строки, обе завершены.
# Менеджер должен получить бонус ОДИН раз.
orders = [
    dup_order("Иван Иванов", variant_date=date(2026, 4, 5), complete=date(2026, 4, 20),
              cal="Марьям", link="https://drive/ivan"),
    dup_order("Иван  Иванов.", complete=date(2026, 4, 25),  # другая запись имени + та же ссылка ниже
              cal="Лена", link="https://drive/ivan"),
]
dashboard._gs.read_all_orders = lambda url, creds, ttl: (orders, ["Апрель_2026"])
res = dashboard.calculate_salary(*APR)
olga = res["employees"].get("Ольга")
olga_total = olga["total"] if olga else 0
check("Дедуп: менеджер получил бонус 1 раз (не 2x300=600)", olga_total == 300,
      f"got {olga_total}")

# Сценарий: разные клиенты — бонус каждому
orders2 = [
    dup_order("Анна", complete=date(2026, 4, 20), cal="Марьям"),
    dup_order("Борис", complete=date(2026, 4, 22), cal="Лена"),
]
dashboard._gs.read_all_orders = lambda url, creds, ttl: (orders2, ["Апрель_2026"])
res2 = dashboard.calculate_salary(*APR)
check("Разные клиенты → бонус каждому (2x300=600)",
      res2["employees"].get("Ольга", {}).get("total", 0) == 600)

# Сценарий: та же ссылка, но имя совсем разное → всё равно дедуп по ссылке
orders3 = [
    dup_order("Иван Иванов", complete=date(2026, 4, 20), cal="Марьям", link="https://drive/x"),
    dup_order("И. Иванов",   complete=date(2026, 4, 25), cal="Лена",   link="https://drive/x"),
]
dashboard._gs.read_all_orders = lambda url, creds, ttl: (orders3, ["Апрель_2026"])
res3 = dashboard.calculate_salary(*APR)
check("Дедуп по ссылке при разных именах → 1 бонус",
      res3["employees"].get("Ольга", {}).get("total", 0) == 300)

# Сценарий: ДОП_ОБУЧЕНИЕ — менеджер не получает бонус даже при завершении
orders4 = [dup_order("Ученик", variant_date=None, complete=date(2026, 4, 20),
                     cal="Лена", tarif="ДОП_ОБУЧЕНИЕ", status=True)]
dashboard._gs.read_all_orders = lambda url, creds, ttl: (orders4, ["Апрель_2026"])
res4 = dashboard.calculate_salary(*APR)
check("ДОП_ОБУЧЕНИЕ → менеджер 0",
      res4["employees"].get("Ольга", {}).get("total", 0) == 0)

print()
print("=" * 60)
print(f"  Итог: прошло {passed}, упало {failed}")
print("=" * 60)
exit(0 if failed == 0 else 1)