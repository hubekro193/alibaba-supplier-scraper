# Alibaba Supplier Scraper 🔍

Automatyczna wyszukiwarka dostawców na Alibaba dla branży torebkowej (ramki metalowe, kopertówki wieczorowe). Wyniki zapisywane do pliku Excel oraz Google Sheets online.

---

## Skrypty

| Plik | Opis |
|------|------|
| `alibaba_dostawcy.py` | Główny skrypt – przeszukuje Alibabę i zapisuje dostawców |
| `usun_duplikaty.py` | Usuwa zduplikowanych dostawców ze statusem „Nowy" z Google Sheets |
| `napraw_oceny.py` | Naprawia błędnie zapisane oceny (np. `46120` → `4.8`) |
| `aktualizuj_statusy.py` | Masowa aktualizacja statusów dostawców w Google Sheets |

---

## Instalacja

```bash
pip install undetected-chromedriver openpyxl gspread google-auth
```

Wymagany **Google Chrome** zainstalowany na komputerze.

---

## Konfiguracja Google Sheets

1. Wejdź na [console.cloud.google.com](https://console.cloud.google.com/)
2. Utwórz projekt → włącz **Google Sheets API** i **Google Drive API**
3. Utwórz **Service Account** → pobierz klucz JSON
4. Zmień nazwę pliku na `google_credentials.json` i umieść w tym samym folderze co skrypty
5. Utwórz arkusz Google Sheets o nazwie **„Dostawcy Alibaba"**
6. Udostępnij arkusz emailowi z pliku JSON (uprawnienia: Edytor)

> ⚠️ Nigdy nie wysyłaj `google_credentials.json` na GitHub!

---

## Użycie

### Wyszukiwanie dostawców
```bash
python alibaba_dostawcy.py
```
Otwiera Chrome, przeszukuje Alibabę według fraz z listy `SEARCH_PHRASES`, zapisuje wyniki do Excela i Google Sheets.

### Usuwanie duplikatów
```bash
python usun_duplikaty.py
```
Usuwa zduplikowane wpisy ze statusem „Nowy". Nie usuwa dostawców z innymi statusami (Zaakceptowany, Odrzucony itp.).

### Naprawa błędnych ocen
```bash
python napraw_oceny.py
```
Google Sheets czasem interpretuje ocenę `4.8` jako datę. Ten skrypt to naprawia.

### Aktualizacja statusów
```bash
python aktualizuj_statusy.py
```
Masowa zmiana statusów dostawców na podstawie listy zahardkodowanej w skrypcie.

---

## Statusy dostawców

| Status | Kolor | Znaczenie |
|--------|-------|-----------|
| Nowy | biały | Nowo znaleziony, nieoceniony |
| W trakcie | żółty | W trakcie sprawdzania |
| Sprawdzony | niebieski | Sprawdzony, bez decyzji |
| Zaakceptowany | zielony | Zatwierdzony do kontaktu |
| Odrzucony | czerwony | Odrzucony |
| Do ponownego sprawdzenia | pomarańczowy | Wymaga ponownej analizy |

---

## Konfiguracja (alibaba_dostawcy.py)

```python
SEARCH_PHRASES = [...]   # frazy do wyszukania
MAX_PAGES      = 5       # stron na frazę (ok. 20 wyników/stronę)
MAX_RUNTIME_MIN = 720    # maks. czas działania (12h)
CAPTCHA_WAIT_S  = 10     # czas oczekiwania na rozwiązanie CAPTCHA
```
