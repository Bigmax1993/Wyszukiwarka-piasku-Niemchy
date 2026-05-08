import csv
import re
import time
import json
import logging
import os
import random
import subprocess
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote_plus, urljoin, urlparse, unquote

import requests
from bs4 import BeautifulSoup  # pyright: ignore[reportMissingModuleSource]
from selenium import webdriver  # pyright: ignore[reportMissingImports]
from selenium.webdriver.common.by import By  # pyright: ignore[reportMissingImports]
from selenium.webdriver.chrome.service import Service  # pyright: ignore[reportMissingImports]
from selenium.common.exceptions import NoSuchElementException, TimeoutException  # pyright: ignore[reportMissingImports]
from selenium.webdriver.support.ui import WebDriverWait  # pyright: ignore[reportMissingImports]
from selenium.webdriver.support import expected_conditions as EC  # pyright: ignore[reportMissingImports]
from webdriver_manager.chrome import ChromeDriverManager  # pyright: ignore[reportMissingImports]

# =========================
# KONFIG
# =========================
OUTPUT_DIR = Path(
    r"C:\Users\kanbu\Documents\Automatyczna wyszukiwarka piasku i wysylka zapytania\Wyniki"
)
OUTPUT_FILE = OUTPUT_DIR / "germany_sand_gravel_contacts.xlsx"
CACHE_FILE = OUTPUT_DIR / "germany_sand_gravel_cache.json"
LOG_FILE = OUTPUT_DIR / "germany_sand_gravel_scraper.log"

SEARCH_TERMS = [
    "kieswerk",
    "kiesgrube",
    "sandgrube",
    "sand und kies werk",
    "sand and gravel quarry",
    "quarry sand gravel",
]
SERPER_DISCOVERY_TERMS = [
    "kieswerk",
    "kieswerke",
    "kiesgrube",
    "kies und sand werk",
    "sand und kies werk",
]

# Niemcy bbox
LAT_MIN, LAT_MAX = 47.3, 54.9
LON_MIN, LON_MAX = 6.0, 14.9
LAT_STEP = 0.9
LON_STEP = 1.1

MAX_SCROLL_ROUNDS = 25
SCROLL_PAUSE = 1.0
HEADLESS_DEFAULT = True
CAPTCHA_CHECK_TIMEOUT = 600  # sekundy
CLOSED_ONLY = False
AUTO_ENRICH_CONTACTS = True
STEP_LOG_WITH_TIMESTAMP = True

SERPER_API_URL = "https://google.serper.dev/search"
SERPER_COUNTRY = "de"
SERPER_LANGUAGE = "de"
SERPER_TIMEOUT = 20
SERPER_DAILY_LIMIT = 120
FORCE_SERPER_LOOKUP = True
DISCOVERY_MODE_DEFAULT = "serper_only"  # serper_only | hybrid | maps_only
SERPER_DISCOVERY_RESULTS_PER_TERM = 20
COUNTRY_RESTRICTION = "DE"

REQUEST_TIMEOUT = 20
MAX_CONTACT_LINKS = 3
HTTP_RETRY_ATTEMPTS = 3
HTTP_BACKOFF_SECONDS = 1.5
WEBSITE_SELENIUM_FALLBACK_ENABLED = True
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.1-flash-lite").strip()
GEMINI_MODELS = os.getenv(
    "GEMINI_MODELS",
    (
        f"{GEMINI_MODEL},"
        "gemini-3-flash,"
        "gemini-2.5-flash-lite,"
        "gemini-2.5-flash,"
        "gemini-2.0-flash-lite,"
        "gemini-2.0-flash"
    ),
).strip()

ENABLE_AUTO_EMAIL = True
EMAIL_SUBJECT_TEMPLATE = "Anfrage: Preis pro Tonne Sand - {company_name}"
EMAIL_SIGNATURE = (
    "MFG Modernerfliesenboden Maksym Swinczak tel +49 15223655399\n"
    "web-site: mfg-fliesen.de"
)
BACKGROUND_ONLY_DEFAULT = True
DAILY_EMAIL_LIMIT = 50
EMAIL_PER_DOMAIN_DAILY_LIMIT = 2
SEND_WINDOW_START_HOUR = 8
SEND_WINDOW_END_HOUR = 18
SUBJECT_VARIANTS = [
    "Anfrage: Preis pro Tonne Sand - {company_name}",
    "Preisanfrage Sand je Tonne - {company_name}",
    "Bitte um Preisinfo: Sand pro Tonne - {company_name}",
]
PROMPT_VARIANTS = [
    "Formeller Stil, konkret, 120-150 Woerter.",
    "Praezise und natuerlich, 110-140 Woerter.",
    "Freundlich-professionell, 120-150 Woerter.",
]
EMAIL_SEND_DELAY_MIN_SECONDS = 22
EMAIL_SEND_DELAY_MAX_SECONDS = 58
EMAIL_SPAMMY_TERMS = [
    "gratis",
    "kostenlos",
    "sonderangebot",
    "rabatt",
    "dringend",
    "jetzt",
    "promo",
    "promotion",
    "werbung",
    "kaufen",
]
SERPER_BAD_DOMAINS = [
    "facebook.com",
    "instagram.com",
    "linkedin.com",
    "gelbeseiten.",
    "wikipedia.org",
]
NON_DE_COUNTRY_HINTS = [
    "austria",
    "osterreich",
    "switzerland",
    "schweiz",
    "poland",
    "polska",
    "czech",
    "cesko",
    "france",
    "italy",
    "belgium",
    "netherlands",
    "nederland",
]
DE_COUNTRY_HINTS = [
    "deutschland",
    "germany",
    "bundesrepublik",
    "de-",
]
SUPPRESSED_EMAIL_LOCALPARTS = {
    "noreply",
    "no-reply",
    "do-not-reply",
    "donotreply",
    "mailer-daemon",
    "postmaster",
}
EXPORT_COLUMNS = [
    "Nazwa zwirowni",
    "adres",
    "kraj zwiazkowy",
    "nr. telefonu",
    "E-mail",
    "Strona internetowa",
]
GERMAN_STATES = [
    "Baden-Wuerttemberg",
    "Bayern",
    "Berlin",
    "Brandenburg",
    "Bremen",
    "Hamburg",
    "Hessen",
    "Mecklenburg-Vorpommern",
    "Niedersachsen",
    "Nordrhein-Westfalen",
    "Rheinland-Pfalz",
    "Saarland",
    "Sachsen",
    "Sachsen-Anhalt",
    "Schleswig-Holstein",
    "Thueringen",
]

SITE_FALLBACK_DRIVER = None
WINDOWS_ENV_FALLBACK_CACHE = {}


class CaptchaRequired(Exception):
    pass


def console_step(message):
    if STEP_LOG_WITH_TIMESTAMP:
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{ts}] [ETAP] {message}", flush=True)
    else:
        print(f"[ETAP] {message}", flush=True)


def is_running_in_jupyter():
    try:
        from IPython.core.getipython import get_ipython

        shell = get_ipython()
        if shell is None:
            return False
        return shell.__class__.__name__ == "ZMQInteractiveShell"
    except Exception:
        return False


def wait_for_user_confirmation(message, jupyter_mode=False):
    print(message)
    if jupyter_mode:
        print("W Jupyter wpisz cokolwiek i naciśnij Enter, aby kontynuować.")
    try:
        input("> ")
    except EOFError:
        # W niektórych trybach uruchomienia notebooka stdin może być niedostępne.
        print("Brak interaktywnego input. Czekam 15 sekund i kontynuuję.")
        time.sleep(15)


def setup_logging():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("germany_sand_gravel_scraper")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )

    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger


def frange(start, stop, step):
    v = start
    while v <= stop:
        yield round(v, 4)
        v += step


def save_csv(rows, path):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    fields = [
        "fraza",
        "nazwa",
        "ocena",
        "liczba_opinii",
        "kategoria",
        "adres",
        "full_address",
        "status",
        "telefon",
        "www",
        "url",
        "lat_center",
        "lon_center",
    ]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


def save_excel(rows, path, logger, cache=None):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    try:
        import pandas as pd  # pyright: ignore[reportMissingImports]

        export_rows = build_export_rows(rows, logger=logger, cache=cache)
        df_contacts = pd.DataFrame(export_rows)
        df_states = pd.DataFrame(build_bundesland_rows(rows))

        with pd.ExcelWriter(path) as writer:
            df_contacts.to_excel(writer, index=False, sheet_name="Kontakty")
            df_states.to_excel(writer, index=False, sheet_name="KrajeZwiazkowe")
    except ImportError as e:
        logger.error(
            "Brak pakietu pandas/openpyxl do zapisu XLSX. Zainstaluj: pip install pandas openpyxl"
        )
        raise e


def extract_bundesland(row):
    explicit = (row.get("bundesland") or "").strip()
    if explicit:
        return explicit
    text = " ".join(
        x for x in [(row.get("full_address") or ""), (row.get("adres") or "")] if x
    ).lower()
    for state in GERMAN_STATES:
        if state.lower() in text:
            return state
    return ""


def sanitize_special_text(value):
    if value is None:
        return ""
    text = str(value)
    text = unquote(text)
    text = re.sub(r"[\x00-\x1F\x7F]", " ", text)
    text = " ".join(text.split())
    return text.strip(" -|")


def enrich_row_with_gemini_cleanup(row, logger, cache):
    gemini_cache = cache.setdefault("gemini_row_enrichment", {})
    cache_key = (
        (row.get("url") or "").strip()
        or f"{(row.get('nazwa') or '').strip()}|{(row.get('www') or '').strip()}"
    )

    address = sanitize_special_text(row.get("full_address") or row.get("adres") or "")
    phone = sanitize_special_text(row.get("phones_found") or row.get("telefon") or "")
    email = sanitize_special_text(row.get("email_target") or row.get("emails_found") or "")
    website = sanitize_special_text(row.get("official_website") or row.get("www") or "")
    company = sanitize_special_text(
        row.get("company_name_clean") or row.get("nazwa") or row.get("company_name_raw") or ""
    )

    if cache_key and cache_key in gemini_cache:
        cached = gemini_cache[cache_key]
        row["company_name_clean"] = cached.get("company_name_clean", company) or company
        row["nazwa"] = row["company_name_clean"]
        row["adres"] = cached.get("address", address) or address
        row["telefon"] = cached.get("phone", phone) or phone
        row["email_target"] = cached.get("email", email) or email
        row["official_website"] = cached.get("website", website) or website
        row["bundesland"] = cached.get("bundesland", "") or row.get("bundesland", "")
        return row

    row["company_name_clean"] = company
    row["nazwa"] = company
    row["adres"] = address
    row["telefon"] = phone
    row["email_target"] = email
    row["official_website"] = website

    api_key = os.getenv("GOOGLE_AI_STUDIO_API_KEY", "").strip()
    if not api_key:
        row["bundesland"] = extract_bundesland(row)
        return row

    states = ", ".join(GERMAN_STATES)
    prompt = (
        "Wyczyść pola rekordu firmy i zwróć wyłącznie JSON. "
        "Usuń śmieciowe znaki specjalne, encje i artefakty OCR. "
        "Wyznacz kraj związkowy Niemiec na podstawie adresu/URL/nazwy. "
        "Jeśli nie da się ustalić pewnie, zwróć pusty bundesland. "
        f"Dozwolone bundesland tylko z listy: {states}. "
        "Format JSON: "
        '{"company_name_clean":"","address":"","phone":"","email":"","website":"","bundesland":""}. '
        f"Dane wejściowe: name={company}; address={address}; phone={phone}; email={email}; website={website}"
    )
    try:
        text, used_model = gemini_generate_text(prompt, logger, api_key, cache=cache)
        console_step(f"Gemini cleanup model: {used_model}")
        match = re.search(r"\{.*\}", text, flags=re.DOTALL)
        parsed = json.loads(match.group(0) if match else text)
        gemini_result = {
            "company_name_clean": sanitize_special_text(
                parsed.get("company_name_clean", company)
            )
            or company,
            "address": sanitize_special_text(parsed.get("address", address)) or address,
            "phone": sanitize_special_text(parsed.get("phone", phone)) or phone,
            "email": sanitize_special_text(parsed.get("email", email)) or email,
            "website": sanitize_special_text(parsed.get("website", website)) or website,
            "bundesland": sanitize_special_text(parsed.get("bundesland", "")),
        }
        if gemini_result["bundesland"] not in GERMAN_STATES:
            gemini_result["bundesland"] = extract_bundesland(row)

        row["company_name_clean"] = gemini_result["company_name_clean"]
        row["nazwa"] = gemini_result["company_name_clean"]
        row["adres"] = gemini_result["address"]
        row["telefon"] = gemini_result["phone"]
        row["email_target"] = gemini_result["email"]
        row["official_website"] = gemini_result["website"]
        row["bundesland"] = gemini_result["bundesland"]
        if cache_key:
            gemini_cache[cache_key] = gemini_result
    except Exception as e:
        logger.warning(f"Gemini cleanup/bundesland fallback dla rekordu ({cache_key}): {e}")
        row["bundesland"] = extract_bundesland(row)

    return row


def build_export_rows(rows, logger=None, cache=None):
    export_rows = []
    for row in rows:
        row = normalize_row_company_name(row)
        if logger is not None and cache is not None:
            row = enrich_row_with_gemini_cleanup(row, logger, cache)
        else:
            row["adres"] = sanitize_special_text(row.get("full_address") or row.get("adres") or "")
            row["telefon"] = sanitize_special_text(row.get("phones_found") or row.get("telefon") or "")
            row["bundesland"] = extract_bundesland(row)
        email = (row.get("email_target") or "").strip()
        if not email:
            emails_found = (row.get("emails_found") or "").strip()
            email = emails_found.split(",", 1)[0].strip() if emails_found else ""

        website = (row.get("official_website") or row.get("www") or "").strip()
        address = (row.get("full_address") or row.get("adres") or "").strip()
        phone = (row.get("phones_found") or row.get("telefon") or "").strip()
        if "," in phone:
            phone = phone.split(",", 1)[0].strip()

        export_rows.append(
            {
                "Nazwa zwirowni": (
                    row.get("company_name_clean") or row.get("nazwa") or ""
                ).strip(),
                "adres": address,
                "kraj zwiazkowy": (row.get("bundesland") or extract_bundesland(row)).strip(),
                "nr. telefonu": phone,
                "E-mail": email,
                "Strona internetowa": website,
            }
        )
    return export_rows


def build_bundesland_rows(rows):
    state_rows = []
    seen = set()
    for row in rows:
        row = normalize_row_company_name(row)
        row_url = (row.get("url") or "").strip()
        row_name = (row.get("company_name_clean") or row.get("nazwa") or "").strip()
        row_state = (row.get("bundesland") or extract_bundesland(row)).strip()
        row_address = sanitize_special_text(row.get("full_address") or row.get("adres") or "")
        row_website = sanitize_special_text(
            row.get("official_website") or row.get("www") or ""
        )

        dedupe_key = row_url or f"{row_name}|{row_state}|{row_address}"
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)

        state_rows.append(
            {
                "Nazwa zwirowni": row_name,
                "kraj zwiazkowy": row_state,
                "adres": row_address,
                "Strona internetowa": row_website,
                "URL": row_url,
            }
        )
    return state_rows


def persist_progress(all_rows, cache, logger, reason=""):
    if reason:
        console_step(f"Persist bieżący stan: {reason}")
    else:
        console_step("Persist bieżący stan")
    save_excel(all_rows, OUTPUT_FILE, logger, cache=cache)
    save_cache(cache, logger)


def load_existing_csv(path, logger):
    rows = []
    seen_urls = set()
    if not path.exists():
        return rows, seen_urls
    logger.info(f"Ładowanie istniejącego CSV: {path}")
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        for r in reader:
            rows.append(r)
            if "url" in r and r["url"]:
                seen_urls.add(r["url"])
    logger.info(f"Wczytano {len(rows)} rekordów z CSV (seen_global={len(seen_urls)})")
    return rows, seen_urls


def load_existing_output(path, logger):
    if path.suffix.lower() != ".xlsx":
        return load_existing_csv(path, logger)

    rows = []
    seen_urls = set()
    if not path.exists():
        return rows, seen_urls

    logger.info(f"Ładowanie istniejącego XLSX: {path}")
    try:
        import pandas as pd  # pyright: ignore[reportMissingImports]

        df = pd.read_excel(path)
        rows = df.fillna("").to_dict(orient="records")
        for r in rows:
            url = r.get("url", "")
            if url:
                seen_urls.add(url)
        logger.info(f"Wczytano {len(rows)} rekordów z XLSX (seen_global={len(seen_urls)})")
    except Exception as e:
        logger.warning(f"Nie udało się wczytać XLSX ({e}) - zaczynam od zera.")
    return rows, seen_urls


def load_cache(logger):
    if not CACHE_FILE.exists():
        logger.info("Brak istniejącego cache JSON - zaczynam od zera.")
        return {
            "places": {},
            "contacts": {},
            "serper": {},
            "serper_daily": {},
            "email_daily": {},
            "email_sent_targets": {},
            "email_domain_daily": {},
            "email_suppression": {},
            "gemini_row_enrichment": {},
            "gemini_disabled_models": {},
        }
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            cache = json.load(f)
        if "places" not in cache:
            cache["places"] = {}
        if "contacts" not in cache:
            cache["contacts"] = {}
        if "serper" not in cache:
            cache["serper"] = {}
        if "serper_daily" not in cache:
            cache["serper_daily"] = {}
        if "email_daily" not in cache:
            cache["email_daily"] = {}
        if "email_sent_targets" not in cache:
            cache["email_sent_targets"] = {}
        if "email_domain_daily" not in cache:
            cache["email_domain_daily"] = {}
        if "email_suppression" not in cache:
            cache["email_suppression"] = {}
        if "gemini_row_enrichment" not in cache:
            cache["gemini_row_enrichment"] = {}
        if "gemini_disabled_models" not in cache:
            cache["gemini_disabled_models"] = {}
        logger.info(f"Wczytano cache JSON: {len(cache['places'])} miejsc.")
        return cache
    except Exception as e:
        logger.warning(f"Nie udało się wczytać cache JSON ({e}) - tworzę nowy.")
        return {
            "places": {},
            "contacts": {},
            "serper": {},
            "serper_daily": {},
            "email_daily": {},
            "email_sent_targets": {},
            "email_domain_daily": {},
            "email_suppression": {},
            "gemini_row_enrichment": {},
            "gemini_disabled_models": {},
        }


def save_cache(cache, logger):
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
        logger.info(
            "Zapisano cache JSON: "
            f"places={len(cache.get('places', {}))}, "
            f"contacts={len(cache.get('contacts', {}))}, "
            f"serper={len(cache.get('serper', {}))}, "
            f"serper_daily={len(cache.get('serper_daily', {}))}, "
            f"email_daily={len(cache.get('email_daily', {}))}, "
            f"gemini_rows={len(cache.get('gemini_row_enrichment', {}))}, "
            f"gemini_disabled_models={len(cache.get('gemini_disabled_models', {}))}"
        )
    except Exception as e:
        logger.error(f"Błąd zapisu cache JSON: {e}")


def build_email_jobs_from_cache_json(logger):
    console_step("Buduję kolejkę maili z pliku cache JSON")
    if not CACHE_FILE.exists():
        logger.info("Brak pliku cache JSON - brak maili do wysyłki.")
        console_step("Brak cache JSON - kolejka maili pusta")
        return []

    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        logger.warning(f"Nie udało się odczytać cache JSON do wysyłki maili: {e}")
        console_step(f"Błąd odczytu cache JSON dla kolejki maili: {e}")
        return []

    contacts = data.get("contacts", {}) if isinstance(data, dict) else {}
    jobs = []
    for place_url, info in contacts.items():
        if not isinstance(info, dict):
            continue
        email_target = (info.get("email_target") or "").strip()
        email_status = (info.get("email_status") or "").strip().lower()
        if not email_target:
            continue
        if email_status == "sent":
            continue
        jobs.append(
            {
                "place_url": place_url,
                "email_target": email_target,
                "company_name": info.get("company_name", "firma"),
                "contact_quality_score": int(info.get("contact_quality_score", 0) or 0),
            }
        )
    logger.info(f"Zbudowano kolejkę maili z JSON: {len(jobs)}")
    console_step(f"Kolejka maili z JSON gotowa: {len(jobs)}")
    return jobs


def get_remaining_daily_email_limit(cache):
    today = date.today().isoformat()
    daily = cache.setdefault("email_daily", {})
    sent_today = int(daily.get(today, 0))
    remaining = max(0, DAILY_EMAIL_LIMIT - sent_today)
    console_step(
        f"Limit EMAIL na {today}: wysłane={sent_today}, pozostało={remaining}, max={DAILY_EMAIL_LIMIT}"
    )
    return today, sent_today, remaining


def increase_daily_email_counter(cache, increment=1):
    today = date.today().isoformat()
    daily = cache.setdefault("email_daily", {})
    daily[today] = int(daily.get(today, 0)) + int(increment)


def get_email_domain(email_target):
    if "@" not in (email_target or ""):
        return ""
    return email_target.split("@", 1)[1].strip().lower()


def get_email_local_part(email_target):
    if "@" not in (email_target or ""):
        return ""
    return email_target.split("@", 1)[0].strip().lower()


def is_email_role_based_or_system(email_target):
    local = get_email_local_part(email_target)
    return local in SUPPRESSED_EMAIL_LOCALPARTS


def is_within_send_window():
    now_hour = datetime.now().hour
    return SEND_WINDOW_START_HOUR <= now_hour < SEND_WINDOW_END_HOUR


def get_domain_remaining_daily_limit(cache, domain):
    today = date.today().isoformat()
    domain_daily = cache.setdefault("email_domain_daily", {}).setdefault(today, {})
    sent_for_domain = int(domain_daily.get(domain, 0))
    remaining = max(0, EMAIL_PER_DOMAIN_DAILY_LIMIT - sent_for_domain)
    return today, sent_for_domain, remaining


def increase_domain_daily_counter(cache, domain, increment=1):
    if not domain:
        return
    today = date.today().isoformat()
    domain_daily = cache.setdefault("email_domain_daily", {}).setdefault(today, {})
    domain_daily[domain] = int(domain_daily.get(domain, 0)) + int(increment)


def is_suppressed_target(cache, email_target):
    suppression = cache.setdefault("email_suppression", {})
    return email_target.lower() in suppression


def mark_suppressed_target(cache, email_target, reason):
    if not email_target:
        return
    suppression = cache.setdefault("email_suppression", {})
    suppression[email_target.lower()] = {
        "reason": reason,
        "date": date.today().isoformat(),
    }


def is_soft_bounce_or_spam_error(error_text):
    lowered = (error_text or "").lower()
    if not lowered:
        return False
    markers = [
        "5.7.1",
        "likely unsolicited",
        "unsolicited mail",
        "message blocked",
        "blocked by policy",
        "mail rejected",
        "spam",
        "temporarily deferred",
        "try again later",
    ]
    return any(marker in lowered for marker in markers)


def sleep_between_emails(logger, target_email):
    delay = random.uniform(EMAIL_SEND_DELAY_MIN_SECONDS, EMAIL_SEND_DELAY_MAX_SECONDS)
    delay = round(delay, 1)
    console_step(f"Przerwa antyspam przed kolejnym mailem ({target_email}): {delay}s")
    logger.info(f"Email send jitter sleep={delay}s target={target_email}")
    time.sleep(delay)


def sanitize_generated_email(subject, body, company_name):
    clean_subject = (subject or "").strip()
    clean_body = (body or "").strip()

    if not clean_subject:
        clean_subject = choose_subject_variant(company_name)
    # Twardy zakaz numerow telefonu w temacie.
    clean_subject = re.sub(r"\+?\d[\d\s()./-]{5,}\d", "", clean_subject)
    clean_subject = clean_subject.replace("!", "").replace("?", "")
    clean_subject = clean_subject[:95].strip(" -")

    lowered_subject = clean_subject.lower()
    if any(term in lowered_subject for term in EMAIL_SPAMMY_TERMS):
        clean_subject = choose_subject_variant(company_name)

    lowered_body = clean_body.lower()
    for term in EMAIL_SPAMMY_TERMS:
        if term in lowered_body:
            clean_body = re.sub(term, "", clean_body, flags=re.IGNORECASE)
    clean_body = re.sub(r"\n{3,}", "\n\n", clean_body).strip()
    return clean_subject, clean_body


def sanitize_sender_name(sender_name):
    text = (sender_name or "").strip()
    if not text:
        return "Maksym Swinczak Firma MFG Modernerfliesenboden GmbH"
    # Usun telefon i wszystko po nim.
    text = re.sub(r"\b(tel|telefon)\b.*$", "", text, flags=re.IGNORECASE).strip()
    # Usun URL i adresy mailowe.
    text = re.sub(r"https?://\S+|\bwww\.\S+|\S+@\S+", "", text, flags=re.IGNORECASE).strip()
    # Usun ciagi wygladajace jak numery telefonow.
    text = re.sub(r"\+?\d[\d\s()./-]{5,}\d", "", text).strip()
    # Ujednolic whitespace.
    text = re.sub(r"\s+", " ", text).strip(" ,;-")
    if not text:
        return "Maksym Swinczak Firma MFG Modernerfliesenboden GmbH"
    # Jesli wystepuje GmbH z dodatkami, obetnij po GmbH.
    m = re.search(r"\bGmbH\b", text, flags=re.IGNORECASE)
    if m:
        text = text[: m.end()].strip(" ,;-")
    return text


def was_email_target_sent_today(cache, email_target):
    if not email_target:
        return False
    today = date.today().isoformat()
    sent_targets = cache.setdefault("email_sent_targets", {}).setdefault(today, [])
    return email_target.lower() in {x.lower() for x in sent_targets}


def mark_email_target_sent_today(cache, email_target):
    if not email_target:
        return
    today = date.today().isoformat()
    sent_targets = cache.setdefault("email_sent_targets", {}).setdefault(today, [])
    lowered = email_target.lower()
    if lowered not in {x.lower() for x in sent_targets}:
        sent_targets.append(email_target)


def get_remaining_daily_serper_limit(cache):
    today = date.today().isoformat()
    daily = cache.setdefault("serper_daily", {})
    used_today = int(daily.get(today, 0))
    remaining = max(0, SERPER_DAILY_LIMIT - used_today)
    console_step(
        f"Limit SERPER na {today}: użyte={used_today}, pozostało={remaining}, max={SERPER_DAILY_LIMIT}"
    )
    return today, used_today, remaining


def increase_daily_serper_counter(cache, increment=1):
    today = date.today().isoformat()
    daily = cache.setdefault("serper_daily", {})
    daily[today] = int(daily.get(today, 0)) + int(increment)


def is_serper_limit_reached_today(cache):
    today, _, remaining = get_remaining_daily_serper_limit(cache)
    flags = cache.setdefault("serper_limit_reached", {})
    if flags.get(today):
        return True
    if remaining <= 0:
        flags[today] = True
        return True
    return False


def mark_serper_limit_reached_today(cache):
    today = date.today().isoformat()
    flags = cache.setdefault("serper_limit_reached", {})
    flags[today] = True


def request_with_retry(method, url, logger, **kwargs):
    last_err = None
    for attempt in range(1, HTTP_RETRY_ATTEMPTS + 1):
        try:
            response = method(url, **kwargs)
            response.raise_for_status()
            return response
        except Exception as e:
            last_err = e
            console_step(
                f"HTTP retry {attempt}/{HTTP_RETRY_ATTEMPTS} dla {url}: {e}"
            )
            if attempt < HTTP_RETRY_ATTEMPTS:
                delay = HTTP_BACKOFF_SECONDS * attempt
                time.sleep(delay)
    logger.warning(f"HTTP nieudane po retry dla {url}: {last_err}")
    if last_err is not None:
        raise last_err
    raise RuntimeError("request_with_retry zakonczone bez odpowiedzi i bez wyjatku")


def choose_subject_variant(company_name):
    idx = abs(hash(company_name or "")) % len(SUBJECT_VARIANTS)
    return SUBJECT_VARIANTS[idx].format(company_name=company_name)


def choose_prompt_variant(company_name):
    idx = abs(hash((company_name or "") + "_prompt")) % len(PROMPT_VARIANTS)
    return PROMPT_VARIANTS[idx]


def score_serper_candidate(link, title="", snippet="", company_name=""):
    text = " ".join([link or "", title or "", snippet or ""]).lower()
    score = 0
    if not link:
        return -999
    if any(bad in text for bad in SERPER_BAD_DOMAINS):
        score -= 120
    if "impressum" in text or "kontakt" in text or "contact" in text:
        score += 20
    if company_name:
        tokens = [t for t in re.split(r"\W+", company_name.lower()) if len(t) >= 4]
        score += sum(12 for t in tokens if t in text)
    if link.startswith("https://"):
        score += 5
    if "/maps/" in link:
        score -= 80
    return score


def is_germany_candidate(link, title="", snippet=""):
    if COUNTRY_RESTRICTION != "DE":
        return True
    text = " ".join([link or "", title or "", snippet or ""]).lower()
    if any(x in text for x in NON_DE_COUNTRY_HINTS):
        return False
    if ".de/" in text or text.endswith(".de"):
        return True
    if any(x in text for x in DE_COUNTRY_HINTS):
        return True
    # Dopuszczamy wynik, jeśli nie ma jawnych wskazówek non-DE.
    return True


def compute_contact_quality_score(row):
    score = 0
    if row.get("email_target"):
        score += 45
    if row.get("phones_found"):
        score += 20
    if row.get("official_website"):
        score += 20
    if row.get("full_address") or row.get("adres"):
        score += 10
    serper_score = int(row.get("serper_source_score", 0) or 0)
    score += min(15, max(0, serper_score // 5))
    return score


def normalize_website(url):
    if not url:
        return ""
    u = url.strip()
    if not u.startswith(("http://", "https://")):
        u = "https://" + u
    return u


def derive_name_from_website(website):
    normalized = normalize_website(website)
    if not normalized:
        return ""
    try:
        host = (urlparse(normalized).netloc or "").lower()
    except Exception:
        return ""
    if host.startswith("www."):
        host = host[4:]
    if not host:
        return ""
    root = host.split(".")[0]
    root = re.sub(r"[^a-z0-9-]+", " ", root, flags=re.IGNORECASE)
    root = root.replace("-", " ")
    root = " ".join(root.split()).strip()
    if not root:
        return ""
    return " ".join(part.capitalize() for part in root.split())


def clean_company_name(name, website=""):
    raw = " ".join((name or "").split()).strip(" -|–—")
    text = raw
    if text:
        # Nazwa firmy zwykle jest przed separatorem, a po nim jest SEO-owy dopisek.
        text = re.split(r"\s+[|–—-]\s+", text, maxsplit=1)[0].strip()
        text = re.sub(
            r"\s+(startseite|homepage|home|wikipedia|firmenliste|hersteller)\s*$",
            "",
            text,
            flags=re.IGNORECASE,
        ).strip()
    if not text:
        text = derive_name_from_website(website)
    return text or "Nieznana firma"


def normalize_row_company_name(row):
    raw_name = (row.get("nazwa") or row.get("company_name_raw") or "").strip()
    website_hint = row.get("official_website") or row.get("www") or row.get("url") or ""
    clean_name = clean_company_name(raw_name, website_hint)
    row["company_name_raw"] = raw_name
    row["company_name_clean"] = clean_name
    row["nazwa"] = clean_name
    return row


def get_serper_api_key():
    return os.getenv("SERPER_API_KEY", "").strip()


def get_env_value(name, default=""):
    val = os.getenv(name)
    if val:
        return val.strip()

    # W Jupyter kernel często nie widzi od razu zmiennych ustawionych przez setx.
    if os.name == "nt":
        if name in WINDOWS_ENV_FALLBACK_CACHE:
            return WINDOWS_ENV_FALLBACK_CACHE[name]
        try:
            cmd = (
                "[Environment]::GetEnvironmentVariable("
                f"'{name}','User'"
                ")"
            )
            out = subprocess.check_output(
                ["powershell", "-NoProfile", "-Command", cmd],
                text=True,
                stderr=subprocess.DEVNULL,
            )
            val = (out or "").strip()
            if val:
                WINDOWS_ENV_FALLBACK_CACHE[name] = val
                return val
        except Exception:
            pass

    return (default or "").strip()


def ensure_ssl_cert_env(logger=None):
    # Ustaw cert bundle automatycznie, gdy systemowe CA w Condzie sa niekompletne.
    if os.getenv("SSL_CERT_FILE"):
        return
    try:
        import certifi

        cert_path = (certifi.where() or "").strip()
        if not cert_path:
            return
        os.environ["SSL_CERT_FILE"] = cert_path
        os.environ.setdefault("REQUESTS_CA_BUNDLE", cert_path)
        msg = f"Ustawiono SSL_CERT_FILE z certifi: {cert_path}"
        if logger:
            logger.info(msg)
        else:
            console_step(msg)
    except Exception as e:
        if logger:
            logger.warning(f"Nie udalo sie ustawic SSL_CERT_FILE z certifi: {e}")


def get_gemini_model_candidates():
    models = []
    for raw in (GEMINI_MODELS or "").split(","):
        model = raw.strip()
        if model and model not in models:
            models.append(model)
    if not models and GEMINI_MODEL:
        models.append(GEMINI_MODEL)
    return models


def get_disabled_gemini_models(cache):
    if cache is None:
        return set()
    disabled = cache.setdefault("gemini_disabled_models", {})
    return {m for m, meta in disabled.items() if isinstance(meta, dict) and meta.get("disabled")}


def mark_gemini_model_disabled(cache, model, reason):
    if cache is None or not model:
        return
    disabled = cache.setdefault("gemini_disabled_models", {})
    disabled[model] = {
        "disabled": True,
        "reason": reason,
        "date": date.today().isoformat(),
    }


def gemini_generate_text(prompt, logger, api_key, cache=None):
    payload = {"contents": [{"parts": [{"text": prompt}]}]}
    last_err = None
    disabled_models = get_disabled_gemini_models(cache)
    for model in get_gemini_model_candidates():
        if model in disabled_models:
            continue
        endpoint = (
            "https://generativelanguage.googleapis.com/v1beta/models/"
            f"{model}:generateContent?key={api_key}"
        )
        try:
            resp = request_with_retry(
                requests.post, endpoint, logger, json=payload, timeout=REQUEST_TIMEOUT
            )
            data = resp.json()
            text = (
                data.get("candidates", [{}])[0]
                .get("content", {})
                .get("parts", [{}])[0]
                .get("text", "")
            )
            if text:
                return text, model
        except Exception as e:
            last_err = e
            logger.warning(f"Gemini model fallback: {model} nieudany ({e})")
            err_text = str(e).lower()
            if "404" in err_text or "not found" in err_text:
                mark_gemini_model_disabled(cache, model, str(e))
                logger.warning(f"Gemini model cached as disabled: {model}")
            time.sleep(0.8)
            continue
    raise last_err or RuntimeError("Brak odpowiedzi z Gemini dla wszystkich modeli.")


def find_emails_in_text(text):
    if not text:
        return []
    found = re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text)
    cleaned = []
    for e in found:
        low = e.lower().strip(".,;:()[]{}<>")
        if low.endswith((".png", ".jpg", ".jpeg", ".svg", ".webp")):
            continue
        if low not in cleaned:
            cleaned.append(low)
    return cleaned


def find_phones_in_text(text):
    if not text:
        return []
    candidates = re.findall(r"(?:\+?\d[\d\s()./-]{6,}\d)", text)
    phones = []
    for c in candidates:
        normalized = " ".join(c.split())
        digits = re.sub(r"\D", "", normalized)
        if len(digits) < 7:
            continue
        if normalized not in phones:
            phones.append(normalized)
    return phones


def normalize_phone_for_compare(phone):
    return re.sub(r"\D", "", phone or "")


def build_company_query_from_row(row):
    """
    Buduje sensowne zapytanie do Serper, nawet gdy nazwa z Maps jest pusta.
    Priorytet: nazwa -> kategoria+adres -> slug z URL.
    """
    name = (row.get("nazwa") or "").strip()
    if name:
        return name

    category = (row.get("kategoria") or "").strip()
    address = (row.get("full_address") or row.get("adres") or "").strip()
    query = " ".join(x for x in [category, address] if x).strip()
    if query:
        return query

    place_url = (row.get("url") or "").strip()
    if "/maps/place/" in place_url:
        try:
            slug = place_url.split("/maps/place/", 1)[1].split("/", 1)[0]
            slug = slug.replace("+", " ").replace("%20", " ").strip()
            if slug:
                return slug
        except Exception:
            pass
    return ""


def reconcile_contact_sources(row, collected):
    """
    Jeśli kontakt z Maps i ze strony jest identyczny, preferuj źródło website (Serper + BS4).
    """
    maps_phone = (row.get("telefon") or "").strip()
    maps_website = normalize_website(row.get("www", ""))
    website = normalize_website(collected.get("website", ""))
    website_phones = collected.get("phones", []) or []

    maps_phone_norm = normalize_phone_for_compare(maps_phone)
    website_phone_norms = {normalize_phone_for_compare(p) for p in website_phones if p}
    same_phone = bool(maps_phone_norm and maps_phone_norm in website_phone_norms)
    same_website = bool(maps_website and website and maps_website == website)

    if same_phone or same_website:
        # Odrzucamy kontakt z Maps jako źródło referencyjne i preferujemy website.
        row["telefon"] = website_phones[0] if website_phones else ""
        row["www"] = website or row.get("www", "")
        row["contact_source"] = "serper_bs4"
        row["maps_contact_rejected"] = "yes"
    else:
        row["contact_source"] = "maps_or_mixed"
        row["maps_contact_rejected"] = "no"

    if not row.get("telefon") and website_phones:
        row["telefon"] = website_phones[0]
    if not row.get("www") and website:
        row["www"] = website

    return row


def search_official_website_with_serper(company_name, address, logger, cache):
    # Serper ma używać wyłącznie nazwy z Google Maps.
    query = (company_name or "").strip()
    if not query:
        console_step("Serper pominięty: puste zapytanie")
        return ""

    serper_cache = cache.setdefault("serper", {})
    if query in serper_cache:
        console_step("Serper cache hit")
        cached = serper_cache[query]
        if isinstance(cached, dict):
            console_step(
                f"Serper cache hit -> url={cached.get('url', '') or 'brak'} score={cached.get('score', 0)}"
            )
            return cached.get("url", "")
        console_step(f"Serper cache hit -> url={cached or 'brak'}")
        return cached

    api_key = get_serper_api_key()
    if not api_key:
        logger.warning("Brak SERPER_API_KEY - pomijam wyszukiwanie Serper.")
        console_step("Serper pominięty: brak SERPER_API_KEY")
        serper_cache[query] = ""
        return ""

    if is_serper_limit_reached_today(cache):
        console_step("Serper globalnie pominięty: limit dzienny już osiągnięty")
        serper_cache[query] = ""
        return ""

    today, used_today, remaining = get_remaining_daily_serper_limit(cache)
    if remaining <= 0:
        logger.warning(
            f"Osiągnięto dzienny limit Serper ({SERPER_DAILY_LIMIT}) dla {today}. "
            "Pomijam kolejne zapytania."
        )
        console_step("Serper pominięty: dzienny limit osiągnięty")
        mark_serper_limit_reached_today(cache)
        serper_cache[query] = ""
        return ""

    headers = {"X-API-KEY": api_key, "Content-Type": "application/json"}
    payload = {"q": query, "gl": SERPER_COUNTRY, "hl": SERPER_LANGUAGE, "num": 5}

    try:
        console_step(f"Serper request: {query}")
        increase_daily_serper_counter(cache, 1)
        resp = request_with_retry(
            requests.post,
            SERPER_API_URL,
            logger,
            headers=headers,
            json=payload,
            timeout=SERPER_TIMEOUT,
        )
        data = resp.json()
    except Exception as e:
        logger.warning(f"Serper błąd dla '{query}': {e}")
        console_step(f"Serper błąd: {e}")
        serper_cache[query] = ""
        return ""

    candidates = []
    for k in ("organic", "places"):
        for item in data.get(k, []) or []:
            link = item.get("link") or item.get("website") or ""
            if link:
                candidates.append(
                    {
                        "link": link,
                        "title": item.get("title", ""),
                        "snippet": item.get("snippet", ""),
                    }
                )

    best_score = 0
    if candidates:
        scored = [
            (
                score_serper_candidate(
                    c.get("link", ""),
                    c.get("title", ""),
                    c.get("snippet", ""),
                    company_name,
                ),
                c,
            )
            for c in candidates[:5]
        ]
        scored.sort(key=lambda x: x[0], reverse=True)
        best_score, best_candidate = scored[0]
        result = normalize_website(best_candidate.get("link", ""))
    else:
        result = ""
    console_step(f"Serper wynik: {result or 'brak'}")
    console_step(f"Serper score najlepszego kandydata: {best_score}")
    serper_cache[query] = {"url": result, "score": best_score}
    return result


def discover_places_with_serper(term, logger, cache):
    query = f"{term} deutschland".strip()
    if not query:
        return []

    api_key = get_serper_api_key()
    if not api_key:
        console_step("Serper discovery pominięte: brak SERPER_API_KEY")
        return []
    if is_serper_limit_reached_today(cache):
        console_step("Serper discovery pominięte: dzienny limit osiągnięty")
        return []

    headers = {"X-API-KEY": api_key, "Content-Type": "application/json"}
    payload = {
        "q": query,
        "gl": SERPER_COUNTRY,
        "hl": SERPER_LANGUAGE,
        "num": SERPER_DISCOVERY_RESULTS_PER_TERM,
    }

    try:
        console_step(f"Serper discovery request: {query}")
        increase_daily_serper_counter(cache, 1)
        resp = request_with_retry(
            requests.post,
            SERPER_API_URL,
            logger,
            headers=headers,
            json=payload,
            timeout=SERPER_TIMEOUT,
        )
        data = resp.json()
    except Exception as e:
        console_step(f"Serper discovery błąd: {e}")
        return []

    rows = []
    seen = set()
    for bucket in ("organic", "places"):
        for item in data.get(bucket, []) or []:
            link = normalize_website(item.get("link") or item.get("website") or "")
            if not link or link in seen:
                continue
            if not is_germany_candidate(
                link, item.get("title", ""), item.get("snippet", "")
            ):
                console_step(f"Pomijam wynik non-DE: {link}")
                continue
            seen.add(link)
            rows.append(
                {
                    "fraza": term,
                    "nazwa": clean_company_name(item.get("title", ""), link),
                    "company_name_raw": (item.get("title") or "").strip(),
                    "company_name_clean": clean_company_name(item.get("title", ""), link),
                    "ocena": "",
                    "liczba_opinii": "",
                    "kategoria": term,
                    "adres": "",
                    "full_address": "",
                    "status": "",
                    "telefon": "",
                    "www": link,
                    "url": link,
                    "lat_center": "",
                    "lon_center": "",
                }
            )
    console_step(f"Serper discovery wyniki dla '{term}': {len(rows)}")
    return rows


def get_site_fallback_driver():
    global SITE_FALLBACK_DRIVER
    if SITE_FALLBACK_DRIVER is None:
        SITE_FALLBACK_DRIVER = build_driver(headless=True)
    return SITE_FALLBACK_DRIVER


def close_site_fallback_driver():
    global SITE_FALLBACK_DRIVER
    if SITE_FALLBACK_DRIVER is not None:
        try:
            SITE_FALLBACK_DRIVER.quit()
        except Exception:
            pass
        SITE_FALLBACK_DRIVER = None


def should_use_selenium_fallback(page_text):
    t = (page_text or "").lower()
    return any(
        marker in t
        for marker in [
            "recaptcha",
            "captcha",
            "unusual traffic",
            "consent",
            "cookie settings",
            "accept all",
            "alle akzeptieren",
            "cookie-richtlinie",
            "datenschutz-präferenzen",
            "datenschutz praferenzen",
        ]
    )


def parse_contacts_from_html(base_url, html):
    soup = BeautifulSoup(html or "", "html.parser")
    page_text = soup.get_text(" ", strip=True)
    emails = find_emails_in_text(page_text)
    phones = find_phones_in_text(page_text)
    contact_urls = []

    for a in soup.find_all("a", href=True):
        href = (a.get("href") or "").strip()
        label = (a.get_text(" ", strip=True) or "").lower()
        if href.startswith("mailto:"):
            email = href.replace("mailto:", "").split("?")[0].strip().lower()
            if email and email not in emails:
                emails.append(email)
        if href.startswith("tel:"):
            phone = href.replace("tel:", "").strip()
            if phone and phone not in phones:
                phones.append(phone)

        if any(k in href.lower() or k in label for k in ["kontakt", "contact", "impressum"]):
            full = urljoin(base_url, href)
            if full not in contact_urls:
                contact_urls.append(full)

    return {
        "emails": emails,
        "phones": phones,
        "contact_urls": contact_urls[:MAX_CONTACT_LINKS],
        "page_text": page_text,
    }


def parse_contacts_from_page(url, logger):
    console_step(f"Pobieram stronę kontaktową: {url}")
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )
    }
    html = ""
    try:
        r = request_with_retry(
            requests.get, url, logger, headers=headers, timeout=REQUEST_TIMEOUT
        )
        html = r.text
    except Exception as e:
        logger.info(f"Nie udało się pobrać strony {url}: {e}")
        console_step(f"Błąd pobrania strony: {url}")

    parsed = parse_contacts_from_html(url, html)
    use_fallback = WEBSITE_SELENIUM_FALLBACK_ENABLED and should_use_selenium_fallback(
        parsed.get("page_text", "")
    )
    if use_fallback:
        console_step("Wykryto cookies/CAPTCHA - fallback Selenium headless")
        try:
            driver = get_site_fallback_driver()
            driver.get(url)
            time.sleep(2)
            dismiss_consent(driver)
            time.sleep(1)
            if is_captcha_page(driver):
                logger.warning(f"CAPTCHA na stronie docelowej (fallback Selenium): {url}")
            else:
                parsed = parse_contacts_from_html(url, driver.page_source)
        except Exception as e:
            logger.warning(f"Selenium fallback nieudany dla {url}: {e}")

    console_step(
        f"Strona przeanalizowana: emaile={len(parsed['emails'])}, telefony={len(parsed['phones'])}, "
        f"linki_kontaktowe={len(parsed['contact_urls'])}"
    )
    return {
        "emails": parsed["emails"],
        "phones": parsed["phones"],
        "contact_urls": parsed["contact_urls"],
    }


def collect_contacts_from_website(website, logger):
    website = normalize_website(website)
    if not website:
        console_step("Brak strony WWW - pomijam zbieranie kontaktów")
        return {"emails": [], "phones": [], "website": "", "source_urls": []}

    console_step(f"Start zbierania kontaktów ze strony: {website}")
    base = parse_contacts_from_page(website, logger)
    emails = list(base["emails"])
    phones = list(base["phones"])
    source_urls = [website]

    for u in base["contact_urls"]:
        console_step(f"Pobieram dodatkową podstronę kontaktową: {u}")
        details = parse_contacts_from_page(u, logger)
        for e in details["emails"]:
            if e not in emails:
                emails.append(e)
        for p in details["phones"]:
            if p not in phones:
                phones.append(p)
        if u not in source_urls:
            source_urls.append(u)

    console_step(
        f"Koniec zbierania kontaktów: emaile={len(emails)}, telefony={len(phones)}, źródła={len(source_urls)}"
    )
    return {"emails": emails, "phones": phones, "website": website, "source_urls": source_urls}


def generate_email_content_gemini(company_name, logger, cache=None):
    console_step(f"Generuję treść maila dla: {company_name}")
    subject_hint = choose_subject_variant(company_name)
    prompt_variant = choose_prompt_variant(company_name)
    api_key = os.getenv("GOOGLE_AI_STUDIO_API_KEY", "").strip()
    if not api_key:
        console_step("Gemini pominięte: brak GOOGLE_AI_STUDIO_API_KEY (fallback)")
        body = (
            "Guten Tag,\n\n"
            f"ich bitte um ein Angebot für Sand von {company_name}.\n"
            "Wie hoch ist Ihr aktueller Preis pro Tonne Sand (zzgl. MwSt.)?\n"
            "Bitte teilen Sie auch die Lieferbedingungen und Mindestabnahmemenge mit.\n\n"
            "Vielen Dank im Voraus.\n\n"
            f"{EMAIL_SIGNATURE}"
        )
        return subject_hint, body

    prompt = (
        "Schreibe eine kurze, professionelle E-Mail auf Deutsch an ein Kieswerk. "
        f"Firmenname: {company_name}. "
        "Absender ist ein Handwerksbetrieb, der regelmaessig Sand fuer Bauprojekte bezieht. "
        "Die E-Mail darf NUR nach dem Preis pro Tonne Sand fragen. "
        "Zusatzfragen sind nur Lieferkosten und Mindestabnahmemenge. "
        "Ergaenze einen kurzen Realwelt-Kontext (bedarf fuer laufende Projekte), ohne werblich zu sein. "
        "Keine anderen Themen, keine Zusammenarbeit, keine Werbung. "
        "Keine Telefonnummern im Betreff. "
        "Vermeide Spam-Trigger: keine Ausrufezeichen, keine Sales-Woerter "
        "(z.B. gratis, kostenlos, rabatt, sonderangebot, promo), keine Grossbuchstabenketten. "
        "Betreff neutral und sachlich, maximal 70 Zeichen. "
        f"Stilvorgabe: {prompt_variant} "
        f"Nutze einen Betreff im Stil von: {subject_hint}. "
        f"Die E-Mail muss am Ende exakt diese Signatur enthalten: {EMAIL_SIGNATURE}. "
        "Gib nur JSON zurueck im Format: "
        '{"subject":"...","body":"..."}'
    )
    try:
        text, used_model = gemini_generate_text(prompt, logger, api_key, cache=cache)
        console_step(f"Gemini email model: {used_model}")
        parsed = json.loads(text)
        subject = parsed.get("subject", "")
        body = parsed.get("body", "")
        subject, body = sanitize_generated_email(subject, body, company_name)
        body = (body or "").rstrip()
        if EMAIL_SIGNATURE not in body:
            body = f"{body}\n\n{EMAIL_SIGNATURE}" if body else EMAIL_SIGNATURE
        console_step("Treść maila wygenerowana przez Gemini")
        return subject, body
    except Exception as e:
        logger.warning(f"Nie udało się wygenerować treści Gemini: {e}")
        console_step(f"Gemini błąd, używam fallback: {e}")
        return (
            subject_hint,
            "Guten Tag,\n\n"
            "wir planen in den naechsten Wochen mehrere laufende Bauvorhaben und "
            "benoetigen dafuer regelmaessig Sand.\n"
            f"Ich moechte daher den aktuellen Preis pro Tonne Sand bei {company_name} erfragen.\n"
            "Bitte senden Sie mir den Preis pro Tonne sowie ggf. Lieferkosten "
            "und die Mindestabnahmemenge.\n\n"
            "Falls diese Anfrage nicht die richtige Stelle erreicht, "
            "waere ich fuer eine kurze Weiterleitung sehr dankbar.\n\n"
            f"{EMAIL_SIGNATURE}",
        )


def send_email_homepl(to_email, subject, body, logger):
    try:
        import yagmail  # pyright: ignore[reportMissingImports]
    except ImportError:
        console_step("Gmail pominięte: brak pakietu yagmail (pip install yagmail)")
        return False, "Brak pakietu yagmail"

    console_step(f"Łączę z Gmail (yagmail) i wysyłam do: {to_email}")
    username = get_env_value("GMAIL_USER")
    password = get_env_value("GMAIL_APP_PASSWORD")
    sender_name = sanitize_sender_name(get_env_value("GMAIL_SENDER_NAME"))

    if not (username and password):
        console_step("Gmail pominięte: brak GMAIL_USER/GMAIL_APP_PASSWORD")
        return False, "Brak GMAIL_USER/GMAIL_APP_PASSWORD"

    try:
        yag = yagmail.SMTP(user=username, password=password)
        contents = [body]
        if sender_name:
            from_header = f"{sender_name} <{username}>"
            yag.send(to=to_email, subject=subject, contents=contents, headers={"From": from_header})
        else:
            yag.send(to=to_email, subject=subject, contents=contents)
        console_step(f"Gmail sukces: {to_email}")
        return True, "wysłano"
    except Exception as e:
        logger.warning(f"Błąd wysyłki maila do {to_email}: {e}")
        console_step(f"Gmail błąd: {to_email} ({e})")
        return False, str(e)


def enrich_row_with_contacts(row, cache, logger):
    row = normalize_row_company_name(row)
    place_url = row.get("url", "")
    contacts_cache = cache.setdefault("contacts", {})

    if place_url in contacts_cache:
        console_step(f"Cache kontaktów hit: {row.get('nazwa', '(brak nazwy)')}")
        cached = contacts_cache[place_url]
        row.update(cached)
        row = normalize_row_company_name(row)
        return row

    maps_website = normalize_website(row.get("www", ""))
    console_step(f"WWW z Maps: {maps_website or 'brak'}")
    website = maps_website
    serper_source_score = 0

    # Domyślnie używamy Serper po nazwie firmy; Maps WWW to fallback.
    if FORCE_SERPER_LOOKUP:
        console_step(
            f"Szukam strony firmowej przez Serper (forced): {row.get('nazwa', '(brak nazwy)')}"
        )
        serper_query = build_company_query_from_row(row)
        serper_website = search_official_website_with_serper(
            serper_query, row.get("full_address") or row.get("adres", ""), logger, cache
        )
        if serper_website:
            website = serper_website
            console_step(f"Wybrano stronę z Serper: {website}")
        else:
            website = maps_website
            console_step("Serper nie zwrócił strony - fallback do Maps WWW")
    elif not website:
        console_step(f"Szukam strony firmowej przez Serper: {row.get('nazwa', '(brak nazwy)')}")
        serper_query = build_company_query_from_row(row)
        website = search_official_website_with_serper(
            serper_query, row.get("full_address") or row.get("adres", ""), logger, cache
        )
        console_step(f"Strona po fallback Serper: {website or 'brak'}")

    serper_query = build_company_query_from_row(row)
    serper_cached = cache.setdefault("serper", {}).get(serper_query, {})
    if isinstance(serper_cached, dict):
        serper_source_score = int(serper_cached.get("score", 0) or 0)
    console_step(f"Serper source score: {serper_source_score}")

    collected = collect_contacts_from_website(website, logger) if website else {
        "emails": [],
        "phones": [],
        "website": "",
        "source_urls": [],
    }
    row = reconcile_contact_sources(row, collected)

    subject = ""
    body = ""
    mail_status = "not_sent"
    target_email = collected["emails"][0] if collected["emails"] else ""

    extra = {
        "company_name": row.get("company_name_clean") or row.get("nazwa", ""),
        "company_name_raw": row.get("company_name_raw", ""),
        "company_name_clean": row.get("company_name_clean", ""),
        "official_website": collected["website"],
        "serper_source_score": serper_source_score,
        "emails_found": ", ".join(collected["emails"]),
        "phones_found": ", ".join(collected["phones"]),
        "contact_sources": ", ".join(collected["source_urls"]),
        "contact_source": row.get("contact_source", "maps_or_mixed"),
        "maps_contact_rejected": row.get("maps_contact_rejected", "no"),
        "email_target": target_email,
        "email_subject": subject,
        "email_body": body,
        "email_status": mail_status,
    }
    extra["contact_quality_score"] = compute_contact_quality_score({**row, **extra})
    row.update(extra)
    console_step(
        f"Kontakt finalny -> source={extra.get('contact_source')} maps_rejected={extra.get('maps_contact_rejected')} "
        f"email={extra.get('email_target') or 'brak'} quality={extra.get('contact_quality_score', 0)}"
    )
    contacts_cache[place_url] = extra
    return row


def parse_card_text(raw):
    parts = [p.strip() for p in raw.split("·") if p.strip()]
    kategoria = parts[0] if len(parts) > 0 else ""
    adres = parts[1] if len(parts) > 1 else ""
    status = parts[2] if len(parts) > 2 else ""
    return kategoria, adres, status


def click_if_exists(driver, by, value):
    try:
        el = driver.find_element(by, value)
        el.click()
        return True
    except Exception:
        return False


def dismiss_consent(driver):
    candidates = [
        (By.XPATH, "//button[contains(., 'Accept all')]"),
        (By.XPATH, "//button[contains(., 'Zaakceptuj wszystko')]"),
        (By.XPATH, "//button[contains(., 'Alle akzeptieren')]"),
        (By.XPATH, "//button[contains(., 'I agree')]"),
    ]
    for by, val in candidates:
        if click_if_exists(driver, by, val):
            time.sleep(1)
            break


def search_url(term, lat, lon, zoom=10.5):
    return f"https://www.google.com/maps/search/{quote_plus(term + ' deutschland')}/@{lat},{lon},{zoom}z"


def build_driver(headless=True):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
    else:
        options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


def is_captcha_page(driver):
    try:
        url = (driver.current_url or "").lower()
    except Exception:
        url = ""

    try:
        title = (driver.title or "").lower()
    except Exception:
        title = ""

    if any(x in url for x in ["/sorry/", "sorry/index", "recaptcha"]):
        return True
    if any(x in title for x in ["unusual traffic", "recaptcha", "robot check"]):
        return True

    captcha_xpaths = [
        "//iframe[contains(@src, 'recaptcha')]",
        "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'i am not a robot')]",
        "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'unusual traffic')]",
        "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'reCAPTCHA')]",
    ]
    for xp in captcha_xpaths:
        try:
            if driver.find_elements(By.XPATH, xp):
                return True
        except Exception:
            continue
    return False


def transfer_cookies(source_driver, target_driver):
    try:
        cookies = source_driver.get_cookies()
    except Exception:
        return
    for cookie in cookies:
        try:
            target_driver.add_cookie(cookie)
        except Exception:
            continue


def handle_captcha(driver, logger, jupyter_mode=False):
    logger.warning("Wykryto CAPTCHA. Przełączam na widoczną przeglądarkę do ręcznego potwierdzenia.")
    print("\n[CAPTCHA] Wykryto CAPTCHA - otwieram przeglądarkę do ręcznego potwierdzenia.")
    current_url = ""
    try:
        current_url = driver.current_url
    except Exception:
        pass

    visible_driver = None
    try:
        visible_driver = build_driver(headless=False)
        visible_driver.get("https://www.google.com")
        transfer_cookies(driver, visible_driver)
        visible_driver.get(current_url or "https://www.google.com/maps")

        wait_for_user_confirmation(
            "[CAPTCHA] Rozwiąż CAPTCHA w otwartym oknie. Po zakończeniu potwierdź tutaj.",
            jupyter_mode=jupyter_mode,
        )

        wait_start = time.time()
        while is_captcha_page(visible_driver):
            if (time.time() - wait_start) > CAPTCHA_CHECK_TIMEOUT:
                raise TimeoutException("Przekroczono czas oczekiwania na rozwiązanie CAPTCHA.")
            wait_for_user_confirmation(
                "[CAPTCHA] Nadal wykrywam CAPTCHA. Dokończ w przeglądarce i potwierdź ponownie.",
                jupyter_mode=jupyter_mode,
            )

        headless_driver = build_driver(headless=True)
        headless_driver.get("https://www.google.com")
        transfer_cookies(visible_driver, headless_driver)
        if current_url:
            headless_driver.get(current_url)
        logger.info("CAPTCHA rozwiązana. Powrót do pracy w tle.")
        print("[CAPTCHA] CAPTCHA rozwiązana. Wracam do trybu tła.\n")
        return headless_driver
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        if visible_driver is not None:
            try:
                visible_driver.quit()
            except Exception:
                pass


def handle_captcha_background(driver, logger):
    """
    Obsługa CAPTCHA bez interakcji użytkownika (tryb tylko w tle).
    """
    logger.warning("Wykryto CAPTCHA w trybie background-only. Reset sesji headless.")
    console_step("CAPTCHA (background-only): reset sesji headless i ponowna próba")
    try:
        driver.quit()
    except Exception:
        pass
    new_driver = build_driver(headless=True)
    new_driver.get("https://www.google.com/maps")
    time.sleep(2)
    return new_driver


def scroll_results_panel(driver):
    panel = None
    try:
        panel = driver.find_element(By.XPATH, "//div[@role='feed']")
    except NoSuchElementException:
        pass

    prev_count = 0
    stable = 0

    for _ in range(MAX_SCROLL_ROUNDS):
        cards = driver.find_elements(By.XPATH, "//a[contains(@href, '/maps/place/')]")
        count_now = len(cards)

        if count_now <= prev_count:
            stable += 1
        else:
            stable = 0
        prev_count = count_now

        if stable >= 4:
            break

        if panel is not None:
            driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", panel)
        else:
            driver.execute_script("window.scrollBy(0, 3000);")

        time.sleep(SCROLL_PAUSE)


def extract_open_status(text: str) -> str:
    if not text:
        return ""

    t = " ".join(text.split()).strip()
    tl = t.lower()

    if "otwarte" in tl:
        return "Otwarte"
    if "tymczasowo zamknięte" in tl or "tymczasowo zamkniete" in tl:
        return "Tymczasowo zamknięte"
    if "zamknięte" in tl or "zamkniete" in tl:
        return "Zamknięte"

    if "geöffnet" in tl or "geoeffnet" in tl:
        return "Geöffnet"
    if "vorübergehend geschlossen" in tl or "voruebergehend geschlossen" in tl:
        return "Vorübergehend geschlossen"
    if "geschlossen" in tl:
        return "Geschlossen"

    if "temporarily closed" in tl:
        return "Temporarily closed"
    if "open" in tl:
        return "Open"
    if "closed" in tl:
        return "Closed"

    return ""


def is_closed_status(status: str) -> bool:
    s = (status or "").strip().lower()
    return any(
        x in s
        for x in [
            "tymczasowo zamknięte",
            "tymczasowo zamkniete",
            "vorübergehend geschlossen",
            "voruebergehend geschlossen",
            "temporarily closed",
        ]
    )


def extract_details_in_new_tab(driver, url):
    phone = ""
    website = ""
    status = ""
    full_address = ""

    base_handle = driver.current_window_handle
    driver.execute_script("window.open(arguments[0], '_blank');", url)
    driver.switch_to.window(driver.window_handles[-1])

    try:
        time.sleep(1.5)
        if is_captcha_page(driver):
            raise CaptchaRequired("CAPTCHA w widoku szczegółów miejsca.")

        tel_links = driver.find_elements(By.XPATH, "//a[starts-with(@href,'tel:')]")
        if tel_links:
            href = tel_links[0].get_attribute("href") or ""
            phone = href.replace("tel:", "").strip()

        for xp in [
            "//a[contains(., 'Website')]",
            "//a[contains(., 'Witryna')]",
            "//a[contains(., 'Webseite')]",
        ]:
            els = driver.find_elements(By.XPATH, xp)
            if els:
                h = els[0].get_attribute("href")
                if h:
                    website = h
                    break

        addr_candidates = driver.find_elements(By.XPATH, "//*[@data-item-id='address']")
        for el in addr_candidates:
            txt = (el.text or "").strip()
            if txt and len(txt) > 6:
                full_address = " ".join(txt.split())
                break

        if not full_address:
            fallback_xpaths = [
                "//*[@aria-label and contains(@aria-label, 'Address')]",
                "//*[@aria-label and contains(@aria-label, 'Adres')]",
                "//*[@aria-label and contains(@aria-label, 'Adresse')]",
            ]
            for xp in fallback_xpaths:
                els = driver.find_elements(By.XPATH, xp)
                for el in els:
                    txt = (el.text or "").strip()
                    if txt and len(txt) > 6:
                        full_address = " ".join(txt.split())
                        break
                if full_address:
                    break

        candidate_texts = []
        status_xpaths = [
            "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZĄĆĘŁŃÓŚŹŻ','abcdefghijklmnopqrstuvwxyząćęłńóśźż'),'otwarte') or contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZĄĆĘŁŃÓŚŹŻ','abcdefghijklmnopqrstuvwxyząćęłńóśźż'),'zamknięte') or contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZĄĆĘŁŃÓŚŹŻ','abcdefghijklmnopqrstuvwxyząćęłńóśźż'),'zamkniete')]",
            "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜẞ','abcdefghijklmnopqrstuvwxyzäöüß'),'geöffnet') or contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜẞ','abcdefghijklmnopqrstuvwxyzäöüß'),'geoeffnet') or contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜẞ','abcdefghijklmnopqrstuvwxyzäöüß'),'geschlossen')]",
            "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'open') or contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'closed')]",
        ]

        for xp in status_xpaths:
            els = driver.find_elements(By.XPATH, xp)
            for el in els[:12]:
                txt = (el.text or "").strip()
                if txt:
                    candidate_texts.append(txt)

        if not candidate_texts:
            panel_candidates = driver.find_elements(By.XPATH, "//div[@role='main'] | //body")
            for el in panel_candidates[:2]:
                txt = (el.text or "").strip()
                if txt:
                    candidate_texts.append(txt)

        for txt in candidate_texts:
            s = extract_open_status(txt)
            if s:
                status = s
                break

    except Exception:
        pass
    finally:
        driver.close()
        driver.switch_to.window(base_handle)

    return phone, website, status, full_address


def get_place_details_with_cache(driver, url, cache, logger):
    places = cache.setdefault("places", {})
    if url in places:
        data = places[url]
        return (
            data.get("phone", ""),
            data.get("website", ""),
            data.get("status", ""),
            data.get("full_address", ""),
        )

    phone, website, status, full_address = extract_details_in_new_tab(driver, url)
    places[url] = {
        "phone": phone,
        "website": website,
        "status": status,
        "full_address": full_address,
    }
    logger.info(f"Dodano do cache: {url}")
    return phone, website, status, full_address


def scrape_term_cell(driver, term, lat, lon, cache, logger):
    logger.info(f"Start komórki: term={term}, lat={lat}, lon={lon}")
    console_step(f"Uruchamiam wyszukiwanie dla frazy='{term}' (lat={lat}, lon={lon})")
    driver.get(search_url(term, lat, lon))
    time.sleep(3)
    if is_captcha_page(driver):
        raise CaptchaRequired("CAPTCHA po wejściu na stronę wyszukiwania.")

    dismiss_consent(driver)

    try:
        WebDriverWait(driver, 25).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '/maps/place/')]"))
        )
    except TimeoutException:
        if is_captcha_page(driver):
            raise CaptchaRequired("CAPTCHA zamiast listy wyników.")
        logger.warning(f"Timeout - brak wyników dla term={term}, lat={lat}, lon={lon}")
        return []

    for xp in [
        "//button[contains(., 'Szukaj w tym obszarze')]",
        "//button[contains(., 'Search this area')]",
        "//button[contains(., 'In diesem Bereich suchen')]",
    ]:
        if click_if_exists(driver, By.XPATH, xp):
            logger.info("Kliknięto 'Szukaj w tym obszarze'")
            time.sleep(2)
            break

    scroll_results_panel(driver)
    console_step(f"Zakończono scroll listy wyników dla frazy='{term}'")

    cards = driver.find_elements(By.XPATH, "//a[contains(@href, '/maps/place/')]")
    rows = []
    seen_local = set()

    logger.info(f"Znaleziono {len(cards)} kart dla term={term}, lat={lat}, lon={lon}")
    console_step(f"Znaleziono {len(cards)} kart wyników dla frazy='{term}'")

    for card in cards:
        href = card.get_attribute("href") or ""
        if not href:
            continue

        place_url = urljoin("https://www.google.com", href)
        if place_url in seen_local:
            continue
        seen_local.add(place_url)

        try:
            raw = card.text.strip()
        except Exception:
            raw = ""

        try:
            h3 = card.find_element(By.XPATH, ".//h3")
            name = h3.text.strip()
        except Exception:
            name = ""

        rating = ""
        reviews = ""
        m_rating = re.search(r"(\d[.,]\d)", raw)
        if m_rating:
            rating = m_rating.group(1).replace(",", ".")

        m_reviews = re.search(r"\(([\d\s.,]+)\)", raw)
        if m_reviews:
            reviews = m_reviews.group(1).replace(" ", "")

        category, address, status_from_list = parse_card_text(raw)
        phone, website, status_from_detail, full_address = get_place_details_with_cache(
            driver, place_url, cache, logger
        )

        status = status_from_detail if status_from_detail else status_from_list

        rows.append(
            {
                "fraza": term,
                "nazwa": clean_company_name(name, website or place_url),
                "company_name_raw": name,
                "company_name_clean": clean_company_name(name, website or place_url),
                "ocena": rating,
                "liczba_opinii": reviews,
                "kategoria": category,
                "adres": address,
                "full_address": full_address,
                "status": status,
                "telefon": phone,
                "www": website,
                "url": place_url,
                "lat_center": lat,
                "lon_center": lon,
            }
        )

    return rows


def run_scraper(
    headless_default=HEADLESS_DEFAULT,
    jupyter_mode=None,
    max_new_rows=None,
    enable_auto_email=None,
    background_only=BACKGROUND_ONLY_DEFAULT,
    dry_run_email=False,
    discovery_mode=DISCOVERY_MODE_DEFAULT,
):
    if jupyter_mode is None:
        jupyter_mode = is_running_in_jupyter()
    if enable_auto_email is None:
        enable_auto_email = ENABLE_AUTO_EMAIL

    logger = setup_logging()
    ensure_ssl_cert_env(logger)
    logger.info("=== START skryptu Google Maps Niemcy (żwirownie/piaskownie) ===")
    logger.info(f"Tryb Jupyter: {'TAK' if jupyter_mode else 'NIE'}")
    logger.info(f"Auto email: {'TAK' if enable_auto_email else 'NIE'}")
    logger.info(f"Background only: {'TAK' if background_only else 'NIE'}")
    logger.info(f"Dry run email: {'TAK' if dry_run_email else 'NIE'}")
    logger.info(f"Discovery mode: {discovery_mode}")
    print("[START] Uruchamiam scraper Google Maps (Niemcy: piaskownie/żwirownie).")
    print(
        f"[TRYB] Jupyter: {'TAK' if jupyter_mode else 'NIE'} | "
        f"Auto email: {'TAK' if enable_auto_email else 'NIE'} | "
        f"Background only: {'TAK' if background_only else 'NIE'} | "
        f"Dry run email: {'TAK' if dry_run_email else 'NIE'} | "
        f"Discovery: {discovery_mode}"
    )
    if max_new_rows is not None:
        logger.info(f"Limit nowych rekordów: {max_new_rows}")
        print(f"[LIMIT] Maksymalnie nowych rekordów w tym uruchomieniu: {max_new_rows}")

    console_step(
        "Konfiguracja run: "
        f"headless_default={headless_default}, "
        f"background_only={background_only}, "
        f"enable_auto_email={enable_auto_email}, "
        f"dry_run_email={dry_run_email}, "
        f"discovery_mode={discovery_mode}"
    )

    driver = None
    if discovery_mode in ("hybrid", "maps_only"):
        effective_headless = True if background_only else headless_default
        driver = build_driver(headless=effective_headless)
        console_step("Sterownik przeglądarki uruchomiony")
    else:
        console_step("Tryb serper_only: bez uruchamiania Selenium/Google Maps")

    all_rows, seen_global = load_existing_output(OUTPUT_FILE, logger)
    cache = load_cache(logger)
    # Gdy eksport XLSX nie zawiera URL (widok biznesowy), deduplikację utrzymujemy przez cache.
    seen_global.update(cache.get("contacts", {}).keys())
    console_step(
        f"Wczytano dane wejściowe: rekordy={len(all_rows)}, unikalne_url={len(seen_global)}"
    )
    total_new_rows = 0
    stop_requested = False

    try:
        if discovery_mode == "serper_only":
            for term in SERPER_DISCOVERY_TERMS:
                if stop_requested:
                    break
                console_step(f"Fraza (serper_only): {term}")
                rows = discover_places_with_serper(term, logger, cache)
                added = 0
                for r in rows:
                    if r["url"] in seen_global:
                        console_step(f"Pomijam duplikat URL: {r.get('url', '')}")
                        continue
                    if AUTO_ENRICH_CONTACTS:
                        console_step(f"Wzbogacam kontakty (serper_only): {r.get('nazwa', '(brak nazwy)')}")
                        r = enrich_row_with_contacts(r, cache, logger)
                    if enable_auto_email and r.get("email_target"):
                        r["email_status"] = "queued"
                        cache.setdefault("contacts", {}).setdefault(r["url"], {})[
                            "email_status"
                        ] = "queued"
                    seen_global.add(r["url"])
                    all_rows.append(r)
                    added += 1
                    total_new_rows += 1
                    if max_new_rows is not None and total_new_rows >= max_new_rows:
                        stop_requested = True
                    persist_progress(
                        all_rows,
                        cache,
                        logger,
                        reason=f"serper_only dodano rekord #{total_new_rows}",
                    )
                print(f"{term}: +{added}")
                persist_progress(all_rows, cache, logger, reason=f"koniec frazy {term}")
        else:
            grid_points = [
                (lat, lon)
                for lat in frange(LAT_MIN, LAT_MAX, LAT_STEP)
                for lon in frange(LON_MIN, LON_MAX, LON_STEP)
            ]
            logger.info(f"Punktów siatki: {len(grid_points)}")
            print(f"Punktów siatki: {len(grid_points)}")
            console_step(f"Rozpoczynam iterację po siatce: {len(grid_points)} punktów")

            for idx, (lat, lon) in enumerate(grid_points, start=1):
                if stop_requested:
                    break
                logger.info(f"=== Komórka {idx}/{len(grid_points)} | lat={lat}, lon={lon} ===")
                print(f"\n=== Komórka {idx}/{len(grid_points)} | lat={lat}, lon={lon} ===")
                console_step(f"Start komórki {idx}/{len(grid_points)}")

                for term in SEARCH_TERMS:
                    if stop_requested:
                        break
                    captcha_retries = 0
                    console_step(f"Fraza: {term}")
                    while True:
                        try:
                            rows = scrape_term_cell(driver, term, lat, lon, cache, logger)
                            added = 0
                            console_step(f"Pobrano {len(rows)} rekordów surowych dla frazy '{term}'")

                            for r in rows:
                                if CLOSED_ONLY and not is_closed_status(r.get("status", "")):
                                    console_step(
                                        f"Pomijam (niezamknięte): {r.get('nazwa', '(brak nazwy)')}"
                                    )
                                    continue
                                if r["url"] in seen_global:
                                    console_step(
                                        f"Pomijam duplikat URL: {r.get('nazwa', '(brak nazwy)')}"
                                    )
                                    continue
                                if AUTO_ENRICH_CONTACTS:
                                    console_step(
                                        f"Wzbogacam kontakty: {r.get('nazwa', '(brak nazwy)')}"
                                    )
                                    r = enrich_row_with_contacts(r, cache, logger)
                                    console_step(
                                        f"Kontakt docelowy: {r.get('email_target', '') or 'brak'}"
                                    )
                                if enable_auto_email and r.get("email_target"):
                                    console_step(
                                        f"Oznaczam do wysyłki końcowej: {r.get('email_target')}"
                                    )
                                    r["email_status"] = "queued"
                                    console_step("Status email: queued")
                                    cache.setdefault("contacts", {}).setdefault(r["url"], {})[
                                        "email_status"
                                    ] = "queued"
                                seen_global.add(r["url"])
                                all_rows.append(r)
                                added += 1
                                total_new_rows += 1
                                console_step(
                                    f"Dodano rekord #{total_new_rows}: {r.get('nazwa', '(brak nazwy)')}"
                                )
                                persist_progress(
                                    all_rows,
                                    cache,
                                    logger,
                                    reason=f"maps dodano rekord #{total_new_rows}",
                                )
                                if max_new_rows is not None and total_new_rows >= max_new_rows:
                                    break

                            logger.info(f"{term}: +{added} rekordów w tej komórce")
                            print(f"{term}: +{added}")
                            persist_progress(
                                all_rows, cache, logger, reason=f"koniec frazy {term}"
                            )
                            if max_new_rows is not None and total_new_rows >= max_new_rows:
                                logger.info("Osiągnięto limit rekordów dla bieżącego uruchomienia.")
                                print("Osiągnięto limit rekordów. Kończę uruchomienie.")
                                console_step("Stop po osiągnięciu limitu rekordów")
                                stop_requested = True
                                break
                            console_step(f"Koniec frazy '{term}'")
                            break

                        except CaptchaRequired as e:
                            captcha_retries += 1
                            logger.warning(f"{term}: {e} (próba {captcha_retries})")
                            console_step(
                                f"CAPTCHA dla '{term}' - próba {captcha_retries}/3"
                            )
                            if captcha_retries > 3:
                                logger.error(f"{term}: zbyt wiele CAPTCHA, pomijam frazę.")
                                print(f"{term}: zbyt wiele CAPTCHA, pomijam.")
                                console_step(f"Pomijam frazę '{term}' po wielu CAPTCHA")
                                break
                            if background_only:
                                driver = handle_captcha_background(driver, logger)
                            else:
                                driver = handle_captcha(
                                    driver, logger, jupyter_mode=jupyter_mode
                                )
                            time.sleep(2)
                            continue
                        except Exception as e:
                            logger.exception(f"{term}: błąd")
                            print(f"{term}: błąd ({e})")
                            console_step(f"Błąd frazy '{term}': {e}")
                            break

        if enable_auto_email:
            console_step("Zapisuję cache przed końcową wysyłką")
            persist_progress(all_rows, cache, logger, reason="przed końcową wysyłką")
            email_jobs = build_email_jobs_from_cache_json(logger)
            email_jobs.sort(key=lambda x: x.get("contact_quality_score", 0), reverse=True)
            console_step("Kolejka maili posortowana wg contact_quality_score")
            for rank, job in enumerate(email_jobs, start=1):
                contacts_cache = cache.setdefault("contacts", {})
                contacts_cache.setdefault(job["place_url"], {})
                contacts_cache[job["place_url"]]["send_priority_rank"] = rank
                for row in all_rows:
                    if row.get("url") == job["place_url"]:
                        row["send_priority_rank"] = rank
                        break
            today, sent_today, remaining = get_remaining_daily_email_limit(cache)
            console_step(
                f"Limit dzienny maili: {DAILY_EMAIL_LIMIT}, wysłane dziś: {sent_today}, pozostało: {remaining}"
            )
            if remaining <= 0:
                console_step("Osiągnięto dzienny limit wysyłki. Pomijam wysyłkę końcową.")
                logger.info("Osiągnięto dzienny limit wysyłki maili.")
                for mail in email_jobs:
                    contacts_cache = cache.setdefault("contacts", {})
                    contacts_cache.setdefault(mail["place_url"], {})
                    contacts_cache[mail["place_url"]]["email_status"] = (
                        f"limit_reached_{today}"
                    )
                    for row in all_rows:
                        if row.get("url") == mail["place_url"]:
                            row["email_status"] = f"limit_reached_{today}"
                            break
                console_step("Zapisuję stan po osiągnięciu limitu email")
                save_excel(all_rows, OUTPUT_FILE, logger, cache=cache)
                save_cache(cache, logger)
                jobs_to_send = []
                jobs_deferred = []
            else:
                if len(email_jobs) > remaining:
                    console_step(
                        f"Kolejka maili ({len(email_jobs)}) przekracza limit dzienny. "
                        f"Wysyłam tylko {remaining}."
                    )
                jobs_to_send = email_jobs[:remaining]
                jobs_deferred = email_jobs[remaining:]

            console_step(
                f"Końcowy etap: wysyłam emaile z danych kontaktowych z JSON ({len(jobs_to_send)})"
            )
            for mail in jobs_to_send:
                target = mail["email_target"]
                domain = get_email_domain(target)
                if was_email_target_sent_today(cache, mail["email_target"]):
                    status = f"duplicate_skipped_{today}"
                    contacts_cache = cache.setdefault("contacts", {})
                    contacts_cache.setdefault(mail["place_url"], {})
                    contacts_cache[mail["place_url"]]["email_status"] = status
                    for row in all_rows:
                        if row.get("url") == mail["place_url"]:
                            row["email_status"] = status
                            break
                    console_step(
                        f"Duplikat dzienny - pomijam wysyłkę do {mail['email_target']}"
                    )
                    continue

                if is_email_role_based_or_system(target):
                    status = f"suppressed_role_based_{today}"
                    contacts_cache = cache.setdefault("contacts", {})
                    contacts_cache.setdefault(mail["place_url"], {})
                    contacts_cache[mail["place_url"]]["email_status"] = status
                    for row in all_rows:
                        if row.get("url") == mail["place_url"]:
                            row["email_status"] = status
                            break
                    console_step(f"Suppress role/system - pomijam {target}")
                    continue

                if is_suppressed_target(cache, target):
                    status = f"suppressed_cached_{today}"
                    contacts_cache = cache.setdefault("contacts", {})
                    contacts_cache.setdefault(mail["place_url"], {})
                    contacts_cache[mail["place_url"]]["email_status"] = status
                    for row in all_rows:
                        if row.get("url") == mail["place_url"]:
                            row["email_status"] = status
                            break
                    console_step(f"Suppress cache - pomijam {target}")
                    continue

                if not is_within_send_window():
                    status = f"deferred_send_window_{today}"
                    contacts_cache = cache.setdefault("contacts", {})
                    contacts_cache.setdefault(mail["place_url"], {})
                    contacts_cache[mail["place_url"]]["email_status"] = status
                    for row in all_rows:
                        if row.get("url") == mail["place_url"]:
                            row["email_status"] = status
                            break
                    console_step("Poza oknem wysyłki - odraczam")
                    continue

                if domain:
                    _, sent_for_domain, remaining_for_domain = get_domain_remaining_daily_limit(
                        cache, domain
                    )
                    if remaining_for_domain <= 0:
                        status = f"deferred_domain_limit_{today}"
                        contacts_cache = cache.setdefault("contacts", {})
                        contacts_cache.setdefault(mail["place_url"], {})
                        contacts_cache[mail["place_url"]]["email_status"] = status
                        for row in all_rows:
                            if row.get("url") == mail["place_url"]:
                                row["email_status"] = status
                                break
                        console_step(
                            f"Limit domeny osiągnięty ({domain}: {sent_for_domain}/{EMAIL_PER_DOMAIN_DAILY_LIMIT})"
                        )
                        continue

                subject, body = generate_email_content_gemini(
                    mail.get("company_name", "firma"), logger, cache=cache
                )
                if dry_run_email:
                    ok, info = True, "dry_run"
                    status = f"dry_run_{today}"
                    console_step(f"DRY RUN: symuluję wysyłkę do {mail['email_target']}")
                else:
                    ok, info = send_email_homepl(
                        mail["email_target"], subject, body, logger
                    )
                    status = "sent" if ok else f"error: {info}"
                    if not ok and is_soft_bounce_or_spam_error(info):
                        status = f"soft_fail_spam_{today}"
                if ok:
                    increase_daily_email_counter(cache, 1)
                    mark_email_target_sent_today(cache, mail["email_target"])
                    increase_domain_daily_counter(cache, domain, 1)
                contacts_cache = cache.setdefault("contacts", {})
                contacts_cache.setdefault(mail["place_url"], {})
                contacts_cache[mail["place_url"]]["email_subject"] = subject
                contacts_cache[mail["place_url"]]["email_body"] = body
                contacts_cache[mail["place_url"]]["email_status"] = status
                if status.startswith("error:"):
                    lowered = status.lower()
                    if "mailbox unavailable" in lowered or "user unknown" in lowered:
                        mark_suppressed_target(cache, target, status)

                for row in all_rows:
                    if row.get("url") == mail["place_url"]:
                        row["email_subject"] = subject
                        row["email_body"] = body
                        row["email_status"] = status
                        break
                console_step(f"Wysyłka do {mail['email_target']}: {status}")
                persist_progress(
                    all_rows,
                    cache,
                    logger,
                    reason=f"status maila {mail['email_target']} -> {status}",
                )
                if not dry_run_email:
                    sleep_between_emails(logger, mail["email_target"])

            for mail in jobs_deferred:
                contacts_cache = cache.setdefault("contacts", {})
                contacts_cache.setdefault(mail["place_url"], {})
                contacts_cache[mail["place_url"]]["email_status"] = f"deferred_{today}"
                for row in all_rows:
                    if row.get("url") == mail["place_url"]:
                        row["email_status"] = f"deferred_{today}"
                        break
            if jobs_deferred:
                console_step(
                    f"Odroczono {len(jobs_deferred)} maili do kolejnego dnia (limit dzienny)."
                )
            persist_progress(all_rows, cache, logger, reason="po końcowej wysyłce")

    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
        close_site_fallback_driver()
        logger.info("Zamknięto przeglądarkę.")
        console_step("Zamknięto sterownik przeglądarki")

    logger.info(f"Gotowe. Zapisano {len(all_rows)} rekordów do: {OUTPUT_FILE}")
    print(f"\nGotowe. Zapisano {len(all_rows)} rekordów do: {OUTPUT_FILE}")
    print("[KONIEC] Scraper zakończył działanie.")


def run_in_jupyter(
    headless=False,
    closed_only=False,
    enable_auto_email=None,
    max_new_rows=None,
    background_only=False,
    dry_run_email=False,
    discovery_mode=DISCOVERY_MODE_DEFAULT,
):
    """
    Główny punkt startowy do uruchamiania z jednej komórki Jupyter.
    """
    if enable_auto_email is None:
        enable_auto_email = ENABLE_AUTO_EMAIL
    console_step(
        "Tryb Jupyter interaktywny: "
        f"headless={headless}, background_only={background_only}, "
        f"enable_auto_email={enable_auto_email}, dry_run_email={dry_run_email}"
    )
    global CLOSED_ONLY
    CLOSED_ONLY = closed_only
    run_scraper(
        headless_default=headless,
        jupyter_mode=True,
        max_new_rows=max_new_rows,
        enable_auto_email=enable_auto_email,
        background_only=background_only,
        dry_run_email=dry_run_email,
        discovery_mode=discovery_mode,
    )


def main():
    """
    Uruchomienie skryptu jako plik .py (terminal lub VS Code Run).
    """
    run_scraper(
        headless_default=HEADLESS_DEFAULT, jupyter_mode=is_running_in_jupyter()
    )


# Brak autostartu: w Jupyter uruchamiaj jawnie:
# run_in_jupyter(headless=False, closed_only=False)
# lub:
# main()

if __name__ == "__main__":
    main()
