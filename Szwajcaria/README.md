# Automatyczna wyszukiwarka piasku i wysylka zapytania (Szwajcaria)

Skrypt do:

- wyszukiwania piaskowni/żwirowni w Szwajcarii (domyślnie Serper, opcjonalnie Google Maps),
- wzbogacania rekordów o kontakty (Serper + `requests` + `BeautifulSoup`),
- zapisu wyników do Excela i JSON cache,
- generowania i wysyłki maili po niemiecku (Gmail przez `yagmail`).

## Pliki projektu

- `switzerland_sand_gravel_scraper.py` - główny skrypt (wariant CH)
- `germany_sand_gravel_scraper.py` - baza współdzielonej logiki
- `pytest.ini` - konfiguracja testów
- `tests/unit` - testy jednostkowe
- `tests/integration` - testy integracyjne
- `tests/regression` - testy regresyjne
- `Wyniki/` - logi, cache i Excel z wynikami

## Wymagania

Python 3.10+ oraz pakiety:

```bash
pip install selenium webdriver-manager requests beautifulsoup4 pandas openpyxl pytest yagmail
```

## Konfiguracja zmiennych środowiskowych

```powershell
setx SERPER_API_KEY "TU_SERPER_API_KEY"
setx GOOGLE_AI_STUDIO_API_KEY "TU_GOOGLE_AI_STUDIO_API_KEY"
setx GMAIL_USER "twoj_mail@gmail.com"
setx GMAIL_APP_PASSWORD "gmail_app_password"
setx GMAIL_SENDER_NAME "Twoja nazwa"
```

## Szybki start

```bash
python "C:\Users\kanbu\Documents\Automatyczna wyszukiwarka piasku i wysylka zapytania\Szwajcaria\switzerland_sand_gravel_scraper.py"
```

## Najważniejsze ustawienia CH

- `SERPER_COUNTRY = "ch"`
- `COUNTRY_RESTRICTION = "CH"`
- zapytania discovery/Maps mają sufiks `schweiz`
- bbox ustawiony pod Szwajcarię
- osobne pliki wynikowe:
  - `switzerland_sand_gravel_contacts.xlsx`
  - `switzerland_sand_gravel_cache.json`
  - `switzerland_sand_gravel_scraper.log`
