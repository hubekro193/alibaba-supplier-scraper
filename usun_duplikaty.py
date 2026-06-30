"""
========================================================
  USUWANIE DUPLIKATÓW – Google Sheets
  Usuwa zduplikowanych dostawców ze statusem "Nowy".

  Zasada:
  - Jeśli firma pojawia się kilka razy i ma różne statusy
    → usuwa tylko kopie ze statusem "Nowy", zostawia resztę
  - Jeśli firma pojawia się kilka razy i wszystkie mają "Nowy"
    → zostawia jeden wiersz (z największą liczbą danych), resztę usuwa
  - Dostawcy z innymi statusami (Zaakceptowany, Odrzucony itd.)
    → nigdy nie są usuwani

  JAK URUCHOMIĆ:
    python usun_duplikaty.py
========================================================
"""

import gspread
from google.oauth2.service_account import Credentials
from collections import defaultdict
import os

# ── Konfiguracja ──
GOOGLE_CREDS_FILE = "google_credentials.json"
GOOGLE_SHEET_NAME = "Dostawcy Alibaba"
STATUS_CHRONIONY  = "Nowy"   # tylko wiersze z tym statusem mogą być usunięte

# Kolumny (0-indexed)
COL_STATUS = 0   # A – Status
COL_NAME   = 2   # C – Nazwa dostawcy

# Ile wierszy usuwamy w jednym batch_update (Google akceptuje do ~1000)
BATCH_SIZE = 500


def batch_delete_rows(ws, row_indices):
    """
    Usuwa wiersze jednym (lub kilkoma) zbiorczymi żądaniami do Sheets API.
    Wysyłamy jedno żądanie HTTP zamiast N osobnych – nie trafia w limit.
    """
    sorted_indices = sorted(set(row_indices), reverse=True)  # od końca do góry

    # Podziel na paczki na wypadek bardzo dużej liczby wierszy
    chunks = [sorted_indices[i:i+BATCH_SIZE] for i in range(0, len(sorted_indices), BATCH_SIZE)]

    deleted = 0
    for chunk in chunks:
        requests_body = []
        for row_idx in chunk:
            requests_body.append({
                "deleteDimension": {
                    "range": {
                        "sheetId": ws.id,
                        "dimension": "ROWS",
                        "startIndex": row_idx - 1,  # Sheets API używa 0-indexed
                        "endIndex":   row_idx,       # exclusive
                    }
                }
            })
        ws.spreadsheet.batch_update({"requests": requests_body})
        deleted += len(chunk)
        print(f"  Usunięto {deleted}/{len(sorted_indices)} wierszy...")

    return deleted


def main():
    print("=" * 55)
    print("  USUWANIE DUPLIKATÓW – Dostawcy Alibaba")
    print("=" * 55)

    # ── Połącz z Google Sheets ──
    if not os.path.exists(GOOGLE_CREDS_FILE):
        print(f"\n❌  Nie znaleziono pliku: '{GOOGLE_CREDS_FILE}'")
        print("   Upewnij się że plik jest w tym samym folderze co skrypt.")
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

    # ── Pobierz wszystkie dane ──
    all_rows = ws.get_all_values()
    if len(all_rows) < 2:
        print("  ℹ️  Arkusz jest pusty lub ma tylko nagłówek.")
        return

    data     = all_rows[1:]   # bez nagłówka
    total_in = len(data)
    print(f"  Wierszy danych: {total_in}")

    # ── Grupuj po nazwie dostawcy ──
    groups = defaultdict(list)
    for i, row in enumerate(data):
        name   = row[COL_NAME].strip().lower() if len(row) > COL_NAME   else ""
        status = row[COL_STATUS].strip()       if len(row) > COL_STATUS else ""
        if not name:
            continue
        sheet_row = i + 2   # +1 nagłówek, +1 bo gspread jest 1-indexed
        filled    = sum(1 for cell in row if cell.strip())
        groups[name].append({
            "sheet_row": sheet_row,
            "status":    status,
            "filled":    filled,
        })

    # ── Zdecyduj które wiersze usunąć ──
    rows_to_delete = []

    for name, entries in groups.items():
        if len(entries) == 1:
            continue

        protected = [e for e in entries if e["status"] != STATUS_CHRONIONY]
        deletable = [e for e in entries if e["status"] == STATUS_CHRONIONY]

        if protected:
            # Jest przynajmniej jeden z innym statusem → usuń wszystkie kopie "Nowy"
            rows_to_delete.extend(e["sheet_row"] for e in deletable)
        else:
            # Wszystkie mają "Nowy" → zostaw jeden (z największą liczbą danych)
            deletable_sorted = sorted(deletable, key=lambda e: e["filled"], reverse=True)
            rows_to_delete.extend(e["sheet_row"] for e in deletable_sorted[1:])

    if not rows_to_delete:
        print("\n  ✅  Brak duplikatów do usunięcia.")
        return

    print(f"\n  Znaleziono {len(rows_to_delete)} wierszy do usunięcia (duplikaty ze statusem '{STATUS_CHRONIONY}').")
    answer = input("  Czy usunąć? (tak / nie): ").strip().lower()
    if answer not in ("tak", "t", "yes", "y"):
        print("  ❌  Anulowano.")
        return

    # ── Usuń zbiorczo (1 żądanie HTTP zamiast N) ──
    print()
    deleted = batch_delete_rows(ws, rows_to_delete)

    remaining = total_in - deleted
    print(f"\n  ✅  Gotowe! Usunięto {deleted} duplikatów.")
    print(f"  Pozostało w arkuszu: {remaining} dostawców.")
    print(f"  🔗 {sh.url}")


if __name__ == "__main__":
    main()
