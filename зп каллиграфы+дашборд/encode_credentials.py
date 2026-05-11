"""
Кодирует credentials.json в base64 для вставки в переменную окружения Render.
Запусти один раз: python encode_credentials.py
Скопируй результат и вставь в Render → Environment → GOOGLE_CREDENTIALS_B64
"""
import base64, os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

creds_file = os.path.join(os.path.dirname(__file__), "credentials.json")

if not os.path.exists(creds_file):
    print("Файл credentials.json не найден!")
    sys.exit(1)

with open(creds_file, "rb") as f:
    encoded = base64.b64encode(f.read()).decode("utf-8")

print("=" * 60)
print("GOOGLE_CREDENTIALS_B64 (скопируй всю строку целиком):")
print("=" * 60)
print(encoded)
print("=" * 60)
print("\nВставь эту строку в Render → Environment Variables")
print("Ключ:    GOOGLE_CREDENTIALS_B64")
print("Значение: (строка выше)")
