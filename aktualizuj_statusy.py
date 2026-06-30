"""
========================================================
  AKTUALIZACJA STATUSÓW – Google Sheets
  Na podstawie zdjęcia arkusza Excel.

  JAK URUCHOMIĆ:
    python aktualizuj_statusy.py
========================================================
"""

import gspread
from google.oauth2.service_account import Credentials
import os

# ── Konfiguracja ──
GOOGLE_CREDS_FILE = "google_credentials.json"
GOOGLE_SHEET_NAME = "Dostawcy Alibaba"

COL_STATUS = 0   # A (0-indexed)
COL_NAME   = 2   # C (0-indexed)

BATCH_SIZE = 500

# ── Statusy z obrazka ──
# Klucz: fragment nazwy (wystarczy unikalny kawałek, bez wielkości liter)
# Wartość: status do ustawienia w Sheets
ZMIANY = {
    # ── Zaakceptowany ──
    "Hefei Cyancraft Info-Tech":                    "Zaakceptowany",
    "Shenzhen Sunrise Plastic Products":            "Zaakceptowany",
    "Guangzhou Ruixing Hardware":                   "Zaakceptowany",
    "Guangzhou Bolian Clothing":                    "Zaakceptowany",
    "Dongguan Ruihong Hardware":                    "Zaakceptowany",
    "Heyuan Zhongcheng Weiye Trading":              "Zaakceptowany",
    "Shenzhen Angel Global Bag":                    "Zaakceptowany",
    "Yingde Mingzhi Hardware":                      "Zaakceptowany",

    # ── Odrzucony ──
    "Ningbo LG Industry":                           "Odrzucony",
    "Guangzhou Ding Xing Jewelry":                  "Odrzucony",
    "Dongguan Fung Hing Hardware Products":         "Odrzucony",
    "Dongguan Baidi Hardware Products":             "Odrzucony",
    "Guangzhou Oudi Hardware":                      "Odrzucony",
    "Chaozhou Fromoce Trading":                     "Odrzucony",
    "Ningbo Hongdao E-Commerce":                    "Odrzucony",
    "Ningbo Lg Textile":                            "Odrzucony",
    "Jiaxing Huablossom Import":                    "Odrzucony",
    "Dongguan Shengxiang Hardware Trading":         "Odrzucony",
    "GURU KIRPA EXPORT HOUSE":                      "Odrzucony",
    "Dongguan Changqin Hardware":                   "Odrzucony",
    "Dongguan Jingcheng Industrial":                "Odrzucony",
    "Shenzhen Meideal Industrial":                  "Odrzucony",
    "Shenzhen Hongshengfeng Hardware":              "Odrzucony",
    "Yiwu Mugu Hardware":                           "Odrzucony",
    "Dongguan City Bewin Gift":                     "Odrzucony",
    "Guangzhou Nice Metal Products":                "Odrzucony",
    "Ningbo Inunion Import":                        "Odrzucony",
    "Guangzhou Xiangxing Hardware":                 "Odrzucony",
    "Guangzhou Tianjun Hardware Products":          "Odrzucony",
}


def find_status(name: str) -> str | None:
    """Szuka pasującego klucza w ZMIANY (dopasowanie przez zawieranie, bez wielkości liter)."""
    name_lower = name.lower().strip()
    for fragment, status in ZMIANY.items():
        if fragment.lower() in name_lower:
            return status
    return None


def main():
    print("=" * 55)
    print("  AKTUALIZACJA STATUSÓW – Dostawcy Alibaba")
    print("=" * 55)

    if not os.path.exists(GOOGLE_CREDS_FILE):
        print(f"\n❌  Nie znaleziono: '{GOOGLE_CREDS_FILE}'")
        return

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
    gc    = gspread.authorize(creds)

    try:
        sh = gc.open(GOOGLE_SHEET_NAME)
        ws = sh.sheet1
        print(f"  ✓ Połączono z: '{GOOGLE_SHEET_NAME}'")
    except gspread.SpreadsheetNotFound:
        print(f"\n❌  Nie znaleziono arkusza '{GOOGLE_SHEET_NAME}'")
        return

    all_rows = ws.get_all_values()
    if len(all_rows) < 2:
        print("  ℹ️  Arkusz jest pusty.")
        return

    data = all_rows[1:]
    print(f"  Wierszy danych: {len(data)}")
    print(f"  Wzorców do dopasowania: {len(ZMIANY)}\n")

    # ── Dopasuj wiersze i zbierz zmiany ──
    updates     = []   # (sheet_row_1based, nowy_status, nazwa)
    not_found   = set(ZMIANY.keys())   # wzorce które nie trafiły w żaden wiersz

    for i, row in enumerate(data):
        if len(row) <= max(COL_STATUS, COL_NAME):
            continue

        name           = row[COL_NAME].strip()
        current_status = row[COL_STATUS].strip()
        new_status     = find_status(name)

        if new_status is None:
            continue   # ten dostawca nie jest na liście zmian

        # Oznacz wzorzec jako znaleziony
        for fragment in ZMIANY:
            if fragment.lower() in name.lower():
                not_found.discard(fragment)

        if current_status == new_status:
            continue   # status już jest prawidłowy – pomijamy

        sheet_row = i + 2
        updates.append((sheet_row, new_status, name, current_status))

    # ── Podsumowanie ──
    if updates:
        print(f"  Wierszy do zaktualizowania: {len(updates)}")
        for row, new_s, name, old_s in updates[:10]:
            print(f"  wiersz {row:>4}: {name[:45]:<45}  {old_s} → {new_s}")
        if len(updates) > 10:
            print(f"  ... i {len(updates) - 10} więcej")
    else:
        print("  ✅  Wszystkie statusy są już aktualne – nic do zmiany.")

    if not_found:
        print(f"\n  ⚠️  Nie znaleziono w arkuszu ({len(not_found)} wzorców):")
        for f in sorted(not_found):
            print(f"    – {f}")

    if not updates:
        return

    answer = input(f"\n  Zaktualizować {len(updates)} wierszy? (tak / nie): ").strip().lower()
    if answer not in ("tak", "t", "yes", "y"):
        print("  ❌  Anulowano.")
        return

    # ── Aktualizuj zbiorczo ──
    cell_list = [
        gspread.Cell(sheet_row, COL_STATUS + 1, new_status)
        for sheet_row, new_status, _, _ in updates
    ]

    print(f"\n  🔄  Aktualizuję {len(cell_list)} komórek...")
    for i in range(0, len(cell_list), BATCH_SIZE):
        chunk = cell_list[i:i + BATCH_SIZE]
        ws.update_cells(chunk, value_input_option="RAW")
        done = min(i + BATCH_SIZE, len(cell_list))
        print(f"  Zaktualizowano {done}/{len(cell_list)}...")

    print(f"\n  ✅  Gotowe! Zaktualizowano {len(updates)} statusów.")
    print(f"  🔗 {sh.url}")


if __name__ == "__main__":
    main()
