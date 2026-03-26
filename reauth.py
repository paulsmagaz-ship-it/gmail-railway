#!/usr/bin/env python3
"""
Повторна авторизація Gmail OAuth.
Запусти: python3 reauth.py
Відкриється браузер → увійди в Google → токен збережеться автоматично.
"""
from google_auth_oauthlib.flow import InstalledAppFlow
import os

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), "credentials.json")
OUTPUT_FILE = "/tmp/new_gmail_token.json"

print("🔐 Запускаю OAuth авторизацію Gmail...")
flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
creds = flow.run_local_server(port=0)

with open(OUTPUT_FILE, "w") as f:
    f.write(creds.to_json())

print(f"✅ Токен збережено: {OUTPUT_FILE}")
print("Тепер виконай:")
print(f'  NEW_TOKEN=$(cat {OUTPUT_FILE}) && cd ~/gmail-railway && ~/bin/railway variables set GOOGLE_TOKEN_JSON="$NEW_TOKEN" && ~/bin/railway up --detach')
