# Automatyczna wyszukiwarka piasku i wysylka zapytania

Skrypt do:

- wyszukiwania piaskowni/żwirowni w Niemczech (domyślnie Serper, opcjonalnie Google Maps),
- wzbogacania rekordów o kontakty (Serper + `requests` + `BeautifulSoup`),
- zapisu wyników do Excela i JSON cache,
- generowania i wysyłki maili po niemiecku (na końcu procesu, z limitami dziennymi).

## Pliki projektu

- `germany_sand_gravel_scraper.py` - główny skrypt
- `pytest.ini` - konfiguracja testów
- `tests/unit` - testy jednostkowe
- `tests/integration` - testy integracyjne
- `tests/regression` - testy regresyjne
- `Wyniki/` - logi, cache i Excel z wynikami

## Wymagania

Python 3.10+ (zalecane) oraz pakiety:

```bash
pip install selenium webdriver-manager requests beautifulsoup4 pandas openpyxl pytest yagmail
```

## Konfiguracja zmiennych środowiskowych

Ustaw w PowerShell (trwale dla użytkownika):

```powershell
setx SERPER_API_KEY "TU_SERPER_API_KEY"
setx GOOGLE_AI_STUDIO_API_KEY "TU_GOOGLE_AI_STUDIO_API_KEY"
setx GMAIL_USER "twoj_mail@gmail.com"
setx GMAIL_APP_PASSWORD "gmail_app_password"
setx GMAIL_SENDER_NAME "Twoja nazwa"
```

Po `setx` zamknij i otwórz nowy terminal / zrestartuj kernel Jupyter.

## Główne ustawienia w kodzie

W `germany_sand_gravel_scraper.py`:

- `OUTPUT_DIR` - katalog wyjściowy (`...\\Wyniki`)
- `ENABLE_AUTO_EMAIL` - czy uruchamiać etap końcowej wysyłki maili
- `BACKGROUND_ONLY_DEFAULT` - tryb wyłącznie headless (bez UI)
- `DISCOVERY_MODE_DEFAULT = "serper_only"` - tryb discovery: `serper_only` / `hybrid` / `maps_only`
- `COUNTRY_RESTRICTION = "DE"` - twarde ograniczenie wyników do Niemiec
- `SERPER_DISCOVERY_TERMS` - lista fraz tylko dla Serper (`kieswerk` i odpowiedniki)
- `DAILY_EMAIL_LIMIT = 50` - limit dziennej wysyłki
- `SERPER_DAILY_LIMIT = 120` - limit dziennych wywołań Serper API
- `FORCE_SERPER_LOOKUP = True` - wymusza lookup strony przez Serper po nazwie z Maps
- `HTTP_RETRY_ATTEMPTS = 3` - liczba prób HTTP dla `requests`
- `HTTP_BACKOFF_SECONDS = 1.5` - bazowy backoff między próbami
- `SUBJECT_VARIANTS` / `PROMPT_VARIANTS` - rotacja A/B tematów i stylu promptu
- `EMAIL_PER_DOMAIN_DAILY_LIMIT = 2` - limit wysyłek dziennie na domenę
- `SEND_WINDOW_START_HOUR` / `SEND_WINDOW_END_HOUR` - okno godzinowe wysyłki

## Kolumny w Excel (wyjście biznesowe)

Plik XLSX zawiera wyłącznie:

- `Nazwa zwirowni`
- `adres`
- `kraj zwiazkowy`
- `nr. telefonu`
- `E-mail`
- `Strona internetowa`

## Jak działa pipeline

1. Skrypt odkrywa rekordy (`serper_only` / `hybrid` / `maps_only`).
2. Dla każdego rekordu zbiera kontakty (strona firmowa + podstrony kontaktowe).
3. Waliduje źródła kontaktu (preferencja `serper_bs4` przy zgodności z Maps).
4. Zapisuje dane na bieżąco do:
   - Excel (`germany_sand_gravel_contacts.xlsx`)
   - JSON cache (`germany_sand_gravel_cache.json`)
5. Na końcu (jeśli `enable_auto_email=True`) czyta kontakty z JSON.
6. Sortuje kolejkę maili po `contact_quality_score` (najlepsze kontakty najpierw).
7. Generuje niemiecką treść maila (zapytanie o cenę za tonę piasku, rotacja promptów A/B).
8. Wysyła maile przez Gmail (yagmail).
9. Aktualizuje statusy wysyłki w Excel i JSON.

## Limity dzienne

- Email:
  - limit: `DAILY_EMAIL_LIMIT = 50`
  - licznik zapisywany w `cache JSON` (`email_daily`)
  - deduplikacja odbiorcy per dzień (ten sam adres nie dostanie 2x)
  - limit per domena (`email_domain_daily`)
  - suppression lista (`email_suppression`)
- Serper API:
  - limit: `SERPER_DAILY_LIMIT = 120`
  - licznik zapisywany w `cache JSON` (`serper_daily`)
  - po osiągnięciu limitu kolejne zapytania Serper są pomijane do końca dnia
  - wybór strony odbywa się z top wyników przez scoring jakości źródła

## Statusy maili

- `not_sent` - rekord przygotowany, bez kolejki
- `queued` - zakolejkowany do etapu końcowej wysyłki
- `sent` - wysłany poprawnie
- `error: ...` - błąd wysyłki
- `deferred_YYYY-MM-DD` - odroczony przez limit dzienny
- `limit_reached_YYYY-MM-DD` - pominięty, bo limit dzienny już wyczerpany
- `duplicate_skipped_YYYY-MM-DD` - pominięty (duplikat odbiorcy w tym samym dniu)
- `dry_run_YYYY-MM-DD` - symulacja wysyłki (bez realnego SMTP)
- `deferred_send_window_YYYY-MM-DD` - poza oknem wysyłki
- `deferred_domain_limit_YYYY-MM-DD` - przekroczony limit domeny
- `suppressed_role_based_YYYY-MM-DD` - adres systemowy/role-based
- `suppressed_cached_YYYY-MM-DD` - adres na liście suppression

## Uruchamianie w Jupyter Lab

Komórka 1:

```python
import sys, importlib
project_dir = r"C:\Users\kanbu\Documents\Automatyczna wyszukiwarka piasku i wysylka zapytania"
if project_dir not in sys.path:
    sys.path.insert(0, project_dir)
import germany_sand_gravel_scraper as app
importlib.reload(app)
print(app.__file__)
```

Komórka 2 (test bez wysyłki):

```python
app.run_in_jupyter(
    headless=True,
    enable_auto_email=False,
    max_new_rows=5,
    background_only=True,
    dry_run_email=True,
    discovery_mode="serper_only"
)
```

Komórka 3 (produkcja z wysyłką):

```python
app.run_in_jupyter(
    headless=True,
    enable_auto_email=True,
    max_new_rows=None,
    background_only=True,
    dry_run_email=False,
    discovery_mode="serper_only"
)
```

## Uruchamianie z terminala

```bash
python "C:\Users\kanbu\Documents\Automatyczna wyszukiwarka piasku i wysylka zapytania\germany_sand_gravel_scraper.py"
```

## Wyniki i logi

W katalogu `Wyniki`:

- `germany_sand_gravel_contacts.xlsx`
- `germany_sand_gravel_cache.json`
- `germany_sand_gravel_scraper.log`

Skrypt drukuje etapy działania w konsoli (`[ETAP] ...`) i zapisuje postęp na bieżąco.

## Testy

Uruchom wszystkie:

```bash
python -m pytest "C:\Users\kanbu\Documents\Automatyczna wyszukiwarka piasku i wysylka zapytania"
```

## Uwagi operacyjne

- Nazwa modułu musi wskazywać właściwy plik z projektu (sprawdzaj `print(app.__file__)` w Jupyter).
- Przy CAPTCHA w trybie background skrypt resetuje sesję headless i ponawia próbę.
- Wysyłka maili odbywa się dopiero po zebraniu kontaktów i odczycie ich z JSON cache.
- Ograniczenia dzienne (email/Serper) są trwałe między uruchomieniami, bo liczniki są zapisane w JSON.
- Zapytania HTTP (`requests.get/post`) mają retry + backoff.
- Wysyłka idzie priorytetowo po `contact_quality_score` (lepsze leady najpierw).
- Kolejka dostaje `send_priority_rank` zapisywany w Excel/JSON.
- Po osiągnięciu limitu Serper kolejne lookupy Serper są globalnie pomijane do końca dnia.
