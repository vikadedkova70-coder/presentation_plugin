
from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow

BASE_DIR = Path(__file__).resolve().parents[1]
CLIENT_SECRET_FILE = BASE_DIR / "secrets" / "client_secret.json"

SCOPES = ["https://www.googleapis.com/auth/drive"]


def main():
    if not CLIENT_SECRET_FILE.exists():
        raise FileNotFoundError(
            f"Не найден OAuth JSON: {CLIENT_SECRET_FILE}"
        )

    print("Открываем авторизацию Google...")

    flow = InstalledAppFlow.from_client_secrets_file(
        str(CLIENT_SECRET_FILE),
        SCOPES,
    )

    creds = flow.run_local_server(
        host="localhost",
        port=0,
        access_type="offline",
        prompt="consent",
    )

    if not creds.refresh_token:
        raise RuntimeError(
            "Google не вернул Refresh Token. "
            "Отзовите доступ приложения в Google Account и повторите авторизацию."
        )

    print("\nАвторизация успешно завершена.")
    print("Скопируйте следующие значения в backend/.env:\n")

    print(f"GOOGLE_OAUTH_CLIENT_ID={creds.client_id}")
    print(f"GOOGLE_OAUTH_CLIENT_SECRET={creds.client_secret}")
    print(f"GOOGLE_OAUTH_REFRESH_TOKEN={creds.refresh_token}")

    print("\nНе публикуйте эти значения и не отправляйте их в Git.")


if __name__ == "__main__":
    main()
