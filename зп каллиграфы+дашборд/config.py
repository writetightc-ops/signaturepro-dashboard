# ─── SignaturePro Dashboard — Настройки ──────────────────────────────────────

# URL Google Таблицы с заказами (менеджеры ведут клиентов)
GOOGLE_SHEETS_URL = "https://docs.google.com/spreadsheets/d/1zVxYAVIXR4cwuknI8lS8wElmRlB2cCJD3ceOogMkWOg/edit"

# URL таблицы руководства (Ставки / История ЗП / Корректировки)
# Содержит:
#   - лист "Ставки"          — тарифные ставки и коэффициенты сотрудников
#   - лист "История_ЗП"      — архив зафиксированных выплат (пишется дашбордом)
#   - лист "Корректировки"   — ручные доплаты/вычеты (заполняет руководство)
#   - лист "Итого_к_выплате" — сводная к выплате (формульная, для бухгалтерии)
MGMT_SHEETS_URL = "https://docs.google.com/spreadsheets/d/1bAjeDKCXtyp_MDlxJsnbyIFl43_l6JwbdXIkxmmpOhA/edit"

# Путь к файлу учётных данных сервисного аккаунта Google
# Как получить:
#   1. console.cloud.google.com → APIs & Services → Enable: Google Sheets API
#   2. Credentials → Create Service Account → Keys → JSON → скачай
#   3. Положи рядом с dashboard.py и укажи имя ниже
#   4. В ОБЕИХ таблицах: Поделиться → добавь email сервисного аккаунта (Редактор)
CREDENTIALS_FILE = "credentials.json"

# Кеш данных из Google Sheets (секунды).
# Увеличь если дашборд нагружает API. Уменьши для более свежих данных.
CACHE_SECONDS = 3600
