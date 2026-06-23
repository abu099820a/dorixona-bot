import os
from dotenv import load_dotenv

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "credentials.json")

PASSWORDS = {
    "snabjenets": os.getenv("SNABJENETS_PASSWORD", "snab123"),
    "shofer":     os.getenv("SHOFER_PASSWORD",     "shofer456"),
    "apteka":     os.getenv("APTEKA_PASSWORD",     "apt789"),
    "sklad":      os.getenv("SKLAD_PASSWORD",      "skl000"),
}

ROLE_NAMES = {
    "snabjenets": "Снабженец",
    "shofer":     "Шофёр",
    "apteka":     "Аптека",
    "sklad":      "Склад",
}
