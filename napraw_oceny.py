"""
========================================================
  NAPRAWA BŁĘDÓW FORMATU OCEN – Google Sheets v2

  Problem: Google Sheets zapisał oceny jako daty,
  teraz widać liczby jak 46120.0 (= numer seryjny daty)

  Jak to działa:
    46120 → 8 kwiecień 2026 → miesiąc=4, dzień=8 → ocena 4.8
    46121 → 9 kwiecień 2026 → miesiąc=4, dzień=9 → ocena 4.9
    46119 → 7 kwiecień 2026 → miesiąc=4, dzień=7 → ocena 4.7

  Obsługuje też stary format: "2026-04-09" → 4.9

  JAK URUCHOMIĆ:
    python napraw_oceny.py
========================================================
"""

import gspread
from google.oauth2.service_account import Credentials
from datetime import date, timedelta
import re
import os

# ── Konfiguracja ──
GOOGLE_CREDS_FILE = "google_credentials.json"
GOOGLE_SHEET_NAME = "Dostawcy Alibaba"
COL_RATING_IDX    = 4      # kolumna E (0-indexed) = "Ocena (★)"

# Epoch dla numerów seryjnych Google Sheets / Excel
SHEETS_EPOCH = date(1899, 12, 30)

BATCH_SIZE = 500


def serial_to_rating(val):
    """
    Konwertuje numer seryjny daty (np. 46120.0) na ocenę (np. 4.8).
    Zwraca float lub None jeśli wartość nie jest błędnym numerem seryjnym.
    """
    raw = str(val).strip().replace(",", ".")   # obsługa "46120,0" (format PL)
    try:
        num = float(raw)
    except ValueError:
        return None

    # Prawidłowe oceny Alibaby: 0.0 – 5.0
    # Numery seryjne dla lat 2020-2030: ok. 43831 – 47483
    # Jeśli liczba > 10 to na pewno nie jest oceną
    if num <= 10:
        return None

    serial = int(num)
    try:
        d = SHEETS_EPOCH + timedelta(days=serial)
        month = d.month
        day   = d.day
        # Alibaba oceny: X.Y gdzie X ∈ 1–5 i Y ∈ 1–9
        if 1 <= month <= 5 and 1 <= day <= 9:
            return float(f"{month}.{day}")
    except (OverflowError, ValueError):
        pass

    return None


def datestr_to_rating(val):
    """
    Konwertuje string z datą (np. "2026-04-09") na ocenę (np. 4.9).
    Zwraca float lub None.
    """
    s = str(val).strip()
    m = re.match(r'^\d{4}-(\d{1,2})-(\d{1,2})$', s)
    if m:
        month, day = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 5 and 1 <= day <= 9:
            return float(f"{month}.{day}")
    return None


def value_to_rating(val):
    """Próbuje obydwu metod – numer seryjny lub string z datą."""
    return serial_to_rating(val) or datestr_to_rating(val)


def main():
    print("=" * 55)
    print("  NAPRAWA OCEN v2 – Dostawcy Alibaba")
    print("=" * 55)

    if not os.path.exists(GOOGLE_CREDS_FILE):
        print(f"\n❌  Nie znaleziono pliku: '{GOOGLE_CREDS_FILE}'")
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
        print("  ℹ️  Arkusz jest pusty lub ma tylko nagłówek.")
        return

    data = all_rows[1:]
    print(f"  Wierszy danych: {len(data)}")

    # ── Znajdź błędne oceny ──
    fixes    = []   # (sheet_row_1based, fixed_float)
    examples = []

    for i, row in enumerate(data):
        if len(row) <= COL_RATING_IDX:
            continue

        raw   = row[COL_RATING_IDX].strip()
        fixed = value_to_rating(raw)

        if fixed is not None:
            sheet_row = i + 2
            fixes.append((sheet_row, fixed))
            if len(examples) < 8:
                examples.append(f"  wiersz {sheet_row:>4}: {raw:>12}  →  {fixed}")

    if not fixes:
        print("\n  ✅  Brak błędnych ocen! Ustawiam tylko format kolumny...")
        _set_number_format(sh, ws)
        print("  ✓  Format kolumny E ustawiony na NUMBER.")
        return

    print(f"\n  Znaleziono {len(fixes)} błędnych ocen:")
    for ex in examples:
        print(ex)
    if len(fixes) > 8:
        print(f"  ... i {len(fixes) - 8} więcej")

    answer = input(f"\n  Naprawić wszystkie {len(fixes)} ocen? (tak / nie): ").strip().lower()
    if answer not in ("tak", "t", "yes", "y"):
        print("  ❌  Anulowano.")
        return

    # ── Aktualizuj komórki paczkami (RAW = bez interpretacji) ──
    print(f"\n  🔧  Naprawiam {len(fixes)} komórek...")

    cell_list = [
        gspread.Cell(row, COL_RATING_IDX + 1, rating)
        for row, rating in fixes
    ]

    for i in range(0, len(cell_list), BATCH_SIZE):
        chunk = cell_list[i:i + BATCH_SIZE]
        ws.update_cells(chunk, value_input_option="RAW")
        done = min(i + BATCH_SIZE, len(cell_list))
        print(f"  Zaktualizowano {done}/{len(cell_list)} komórek...")

    # ── Ustaw format NUMBER na kolumnie E – żeby nie powtórzyło się ──
    print("\n  🎨  Ustawiam format NUMBER na kolumnie Ocena (★)...")
    _set_number_format(sh, ws)

    print(f"\n  ✅  Gotowe! Naprawiono {len(fixes)} ocen.")
    print(f"  🔗 {sh.url}")


def _set_number_format(sh, ws):
    """Wymusza format liczbowy 0.0 na kolumnie E (Ocena ★)."""
    try:
        sh.batch_update({"requests": [{
            "repeatCell": {
                "range": {
                    "sheetId": sh.sheet1.id,
                    "startRowIndex": 1,
                    "endRowIndex": 10000,
                    "startColumnIndex": COL_RATING_IDX,
                    "endColumnIndex": COL_RATING_IDX + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {"type": "NUMBER", "pattern": "0.0"},
                    }
                },
                "fields": "userEnteredFormat.numberFormat",
            }
        }]})
    except Exception as e:
        print(f"   ⚠️  Format (niekrytyczne): {e}")


if __name__ == "__main__":
    main()
