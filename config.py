# ─── SignaturePro Dashboard — Настройки ──────────────────────────────────────
#
# Есть два режима работы:
#
#   РЕЖИМ 1 (по умолчанию) — Локальный Excel
#     Дашборд читает SignaturePro_Заказы_NEW.xlsx из этой папки.
#     Заказы редактируются в Excel и сохраняются локально.
#
#   РЕЖИМ 2 — Google Таблица (облако)
#     Дашборд читает и пишет прямо в Google Sheets.
#     Данные актуальны у всех сотрудников в реальном времени.
#     Заказы редактируются в браузере, в Google Таблицах.
#
# Чтобы включить РЕЖИМ 2, заполни GOOGLE_SHEETS_URL и настрой credentials.json
# ─────────────────────────────────────────────────────────────────────────────

# Вставь URL своей Google Таблицы (или оставь пустым для режима Excel)
GOOGLE_SHEETS_URL = "https://docs.google.com/spreadsheets/d/1zVxYAVIXR4cwuknI8lS8wElmRlB2cCJD3ceOogMkWOg/edit"
# Пример:
# GOOGLE_SHEETS_URL = "https://docs.google.com/spreadsheets/d/1ABC123xyz.../edit"

# Путь к файлу учётных данных сервисного аккаунта Google
# Как получить:
#   1. Зайди на https://console.cloud.google.com
#   2. Создай проект → APIs & Services → Enable APIs → Google Sheets API
#   3. Credentials → Create Credentials → Service Account
#   4. Открой сервисный аккаунт → Keys → Add Key → JSON → скачай файл
#   5. Положи скачанный файл рядом с dashboard.py и укажи его имя ниже
#   6. В Google Таблице нажми «Поделиться» и добавь email сервисного аккаунта
#      (он выглядит как something@project-id.iam.gserviceaccount.com)
#      с правами «Редактор»
CREDENTIALS_FILE = "credentials.json"

# Кеш данных из Google Sheets (секунды).
# Данные не перечитываются заново, если прошло меньше этого времени.
# Увеличь, если дашборд нагружает API. Уменьши для более свежих данных.
CACHE_SECONDS = 3600
