import logging
import sys
from types import SimpleNamespace

import germany_sand_gravel_scraper as scraper


class DummyResponse:
    def __init__(self, text="", json_data=None):
        self.text = text
        self._json_data = json_data or {}

    def raise_for_status(self):
        return None

    def json(self):
        return self._json_data


def test_search_official_website_with_serper_uses_api_and_cache(monkeypatch):
    logger = logging.getLogger("test")
    cache = {"serper": {}, "serper_daily": {}}
    calls = {"count": 0}

    def fake_post(url, headers=None, json=None, timeout=None):
        calls["count"] += 1
        return DummyResponse(
            json_data={"organic": [{"link": "https://firma-kies.de/kontakt"}]}
        )

    monkeypatch.setenv("SERPER_API_KEY", "x")
    monkeypatch.setattr(scraper.requests, "post", fake_post)

    first = scraper.search_official_website_with_serper("Firma", "Berlin", logger, cache)
    second = scraper.search_official_website_with_serper("Firma", "Berlin", logger, cache)

    assert first == "https://firma-kies.de/kontakt"
    assert second == first
    assert calls["count"] == 1
    assert isinstance(cache["serper"]["Firma"], dict)
    assert "score" in cache["serper"]["Firma"]


def test_collect_contacts_from_website_merges_base_and_contact_pages(monkeypatch):
    logger = logging.getLogger("test")

    html_main = """
    <html><body>
      Kontakt: info@kies.de
      <a href="tel:+49 30 123456">Telefon</a>
      <a href="/kontakt">Kontakt</a>
    </body></html>
    """
    html_contact = """
    <html><body>
      Vertrieb: sales@kies.de
      <a href="tel:+49 30 999999">Tel 2</a>
    </body></html>
    """

    def fake_get(url, headers=None, timeout=None):
        if url.endswith("/kontakt"):
            return DummyResponse(text=html_contact)
        return DummyResponse(text=html_main)

    monkeypatch.setattr(scraper.requests, "get", fake_get)
    got = scraper.collect_contacts_from_website("https://kies.de", logger)

    assert "info@kies.de" in got["emails"]
    assert "sales@kies.de" in got["emails"]
    assert any("+49 30 123456" in p for p in got["phones"])
    assert any("+49 30 999999" in p for p in got["phones"])
    assert "https://kies.de/kontakt" in got["source_urls"]


def test_send_email_homepl_success(monkeypatch):
    logger = logging.getLogger("test")

    calls = {"sent": False}

    class FakeYagSMTP:
        def __init__(self, user, password):
            self.user = user
            self.password = password

        def send(self, to, subject, contents, headers=None):
            calls["sent"] = True
            assert to == "client@kies.de"
            assert subject == "Subj"
            assert "Body text" in contents[0]

    monkeypatch.setenv("GMAIL_USER", "user@gmail.com")
    monkeypatch.setenv("GMAIL_APP_PASSWORD", "secret")
    monkeypatch.setitem(sys.modules, "yagmail", SimpleNamespace(SMTP=FakeYagSMTP))

    ok, info = scraper.send_email_homepl(
        "client@kies.de", "Subj", "Body text", logger
    )
    assert ok is True
    assert info == "wysłano"
    assert calls["sent"] is True


def test_run_scraper_with_limit_and_email_enabled(monkeypatch):
    logger = logging.getLogger("test")
    state = {"excel_saves": 0, "cache_saves": 0, "emails": 0}

    monkeypatch.setattr(scraper, "setup_logging", lambda: logger)
    monkeypatch.setattr(scraper, "build_driver", lambda headless=True: SimpleNamespace(quit=lambda: None))
    monkeypatch.setattr(scraper, "load_existing_output", lambda path, _logger: ([], set()))
    monkeypatch.setattr(
        scraper,
        "load_cache",
        lambda _logger: {
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
        },
    )
    monkeypatch.setattr(scraper, "frange", lambda start, stop, step: [50.0])
    monkeypatch.setattr(
        scraper,
        "scrape_term_cell",
        lambda driver, term, lat, lon, cache, _logger: [
            {
                "fraza": term,
                "nazwa": "Kies GmbH",
                "ocena": "",
                "liczba_opinii": "",
                "kategoria": "Kieswerk",
                "adres": "Berlin",
                "full_address": "Berlin",
                "status": "",
                "telefon": "",
                "www": "https://kies.de",
                "url": f"https://maps.google.com/{term}",
                "lat_center": lat,
                "lon_center": lon,
            }
        ],
    )
    monkeypatch.setattr(
        scraper,
        "enrich_row_with_contacts",
        lambda row, cache, _logger: {
            **row,
            "email_target": "kontakt@kies.de",
            "contact_quality_score": 90,
        },
    )
    monkeypatch.setattr(scraper, "is_within_send_window", lambda: True)
    monkeypatch.setattr(scraper, "build_email_jobs_from_cache_json", lambda _logger: [])
    monkeypatch.setattr(
        scraper,
        "generate_email_content_gemini",
        lambda company_name, _logger, cache=None: ("Betreff", "Inhalt"),
    )
    monkeypatch.setattr(
        scraper, "send_email_homepl", lambda to, sub, body, _logger: (True, "wysłano")
    )
    monkeypatch.setattr(
        scraper,
        "save_excel",
        lambda rows, path, _logger, cache=None: state.__setitem__(
            "excel_saves", state["excel_saves"] + 1
        ),
    )
    monkeypatch.setattr(
        scraper,
        "save_cache",
        lambda cache, _logger: state.__setitem__("cache_saves", state["cache_saves"] + 1),
    )

    scraper.run_scraper(
        headless_default=True,
        jupyter_mode=True,
        max_new_rows=1,
        enable_auto_email=True,
    )

    assert state["excel_saves"] >= 1
    assert state["cache_saves"] >= 1


def test_run_scraper_dry_run_sets_status_without_real_smtp(monkeypatch):
    logger = logging.getLogger("test")
    state = {"rows": []}

    monkeypatch.setattr(scraper, "setup_logging", lambda: logger)
    monkeypatch.setattr(
        scraper, "build_driver", lambda headless=True: SimpleNamespace(quit=lambda: None)
    )
    monkeypatch.setattr(scraper, "load_existing_output", lambda path, _logger: ([], set()))
    monkeypatch.setattr(
        scraper,
        "load_cache",
        lambda _logger: {
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
        },
    )
    monkeypatch.setattr(scraper, "frange", lambda start, stop, step: [50.0])
    monkeypatch.setattr(
        scraper,
        "scrape_term_cell",
        lambda driver, term, lat, lon, cache, _logger: [
            {
                "fraza": term,
                "nazwa": "Kies GmbH",
                "ocena": "",
                "liczba_opinii": "",
                "kategoria": "Kieswerk",
                "adres": "Berlin",
                "full_address": "Berlin",
                "status": "",
                "telefon": "",
                "www": "https://kies.de",
                "url": "https://maps.google.com/place",
                "lat_center": lat,
                "lon_center": lon,
            }
        ],
    )
    monkeypatch.setattr(
        scraper,
        "enrich_row_with_contacts",
        lambda row, cache, _logger: {
            **row,
            "email_target": "kontakt@kies.de",
            "contact_quality_score": 99,
        },
    )
    monkeypatch.setattr(scraper, "is_within_send_window", lambda: True)
    monkeypatch.setattr(
        scraper,
        "build_email_jobs_from_cache_json",
        lambda _logger: [
            {
                "place_url": "https://maps.google.com/place",
                "email_target": "kontakt@kies.de",
                "company_name": "Kies GmbH",
                "contact_quality_score": 99,
            }
        ],
    )
    monkeypatch.setattr(
        scraper,
        "generate_email_content_gemini",
        lambda company_name, _logger, cache=None: ("Betreff", "Body"),
    )
    monkeypatch.setattr(
        scraper,
        "save_excel",
        lambda rows, path, _logger, cache=None: state.__setitem__("rows", list(rows)),
    )
    monkeypatch.setattr(scraper, "save_cache", lambda cache, _logger: None)

    def fail_send(*args, **kwargs):
        raise AssertionError("SMTP should not be called in dry_run mode")

    monkeypatch.setattr(scraper, "send_email_homepl", fail_send)

    scraper.run_scraper(
        headless_default=True,
        jupyter_mode=True,
        max_new_rows=1,
        enable_auto_email=True,
        dry_run_email=True,
        discovery_mode="maps_only",
    )

    assert state["rows"]
    assert any(str(r.get("email_status", "")).startswith("dry_run_") for r in state["rows"])
