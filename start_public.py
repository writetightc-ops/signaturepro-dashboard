"""
SignaturePro — Публичный дашборд ЗП
Запуск: python start_public.py

Создаёт публичную HTTPS-ссылку через ngrok.
Поделись ей с командой — откроется в любом браузере без VPN и установок.
"""

# ─── Вставь свой ngrok authtoken ──────────────────────────────────────────────
# Как получить (бесплатно, 1 минута):
#   1. Зайди на https://ngrok.com → Sign up (можно через Google)
#   2. После входа: https://dashboard.ngrok.com/get-started/your-authtoken
#   3. Скопируй токен и вставь сюда:
NGROK_AUTHTOKEN = ""
# Пример: NGROK_AUTHTOKEN = "2abc123xyz_AbCdEfGhIjK..."
# ─────────────────────────────────────────────────────────────────────────────

import sys, os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

try:
    from pyngrok import ngrok, conf
except ImportError:
    print("Установи pyngrok: pip install pyngrok")
    sys.exit(1)

PORT = 5000

print("=" * 58)
print("  SignaturePro — Публичный дашборд ЗП")
print("=" * 58)

# Устанавливаем токен
if NGROK_AUTHTOKEN:
    conf.get_default().auth_token = NGROK_AUTHTOKEN
else:
    # Пробуем без токена (работает если уже авторизован через ngrok authtoken)
    pass

try:
    tunnel = ngrok.connect(PORT, bind_tls=True)
    public_url = tunnel.public_url

    print(f"""
  ПУБЛИЧНАЯ ССЫЛКА (поделись с командой):
  {public_url}

  Локальный адрес: http://localhost:{PORT}
  Ctrl+C — остановить сервер
""")
    print("=" * 58 + "\n")

except Exception as e:
    err = str(e)
    print(f"\n  Ошибка ngrok: {err}")
    if "authtoken" in err.lower() or "auth" in err.lower():
        print("""
  Нужен authtoken ngrok (бесплатно):
    1. https://ngrok.com → Sign up
    2. https://dashboard.ngrok.com/get-started/your-authtoken
    3. Вставь токен в NGROK_AUTHTOKEN в этом файле
""")
    sys.exit(1)

# Запускаем Flask
from dashboard import app
app.run(debug=False, host="0.0.0.0", port=PORT)
