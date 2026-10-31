"""Одноразово получает refresh-токен Google для .env.

Запуск (из корня проекта, с включённым .venv):
    pip install google-auth-oauthlib
    python scripts/get_refresh_token.py
Нужен файл secrets/client_secret.json (OAuth-клиент типа «Приложение для ПК»).
"""
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/drive"]

flow = InstalledAppFlow.from_client_secrets_file("secrets/client_secret.json", SCOPES)
creds = flow.run_local_server(port=0, access_type="offline", prompt="consent")

print("\n=== Скопируйте эти строки в файл .env ===")
print(f"GOOGLE_OAUTH_CLIENT_ID={creds.client_id}")
print(f"GOOGLE_OAUTH_CLIENT_SECRET={creds.client_secret}")
print(f"GOOGLE_OAUTH_REFRESH_TOKEN={creds.refresh_token}")
