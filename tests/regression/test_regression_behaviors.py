import json
import logging

import germany_sand_gravel_scraper as scraper


def test_regression_gemini_fallback_always_contains_signature(monkeypatch):
    logger = logging.getLogger("test")

    monkeypatch.delenv("GOOGLE_AI_STUDIO_API_KEY", raising=False)
    subject, body = scraper.generate_email_content_gemini("Kies Nord", logger)

    assert "Sand" in subject
    assert scraper.EMAIL_SIGNATURE in body
    assert "Guten Tag" in body


def test_regression_gemini_response_without_signature_gets_appended(monkeypatch):
    logger = logging.getLogger("test")

    class DummyResponse:
        def raise_for_status(self):
            return None

        def json(self):
            payload = {"subject": "Preis", "body": "Guten Tag,\nPreis bitte."}
            return {"candidates": [{"content": {"parts": [{"text": json.dumps(payload)}]}}]}

    monkeypatch.setenv("GOOGLE_AI_STUDIO_API_KEY", "x")
    monkeypatch.setattr(scraper.requests, "post", lambda *args, **kwargs: DummyResponse())

    _, body = scraper.generate_email_content_gemini("Kies Sued", logger)
    assert scraper.EMAIL_SIGNATURE in body


def test_regression_load_cache_initializes_missing_sections(tmp_path, monkeypatch):
    cache_path = tmp_path / "cache.json"
    cache_path.write_text('{"places": {"a": 1}}', encoding="utf-8")
    logger = logging.getLogger("test")

    monkeypatch.setattr(scraper, "CACHE_FILE", cache_path)
    loaded = scraper.load_cache(logger)

    assert "places" in loaded
    assert "contacts" in loaded
    assert "serper" in loaded
    assert "serper_daily" in loaded
    assert "email_daily" in loaded
    assert "email_sent_targets" in loaded
    assert "email_domain_daily" in loaded
    assert "email_suppression" in loaded
    assert "gemini_row_enrichment" in loaded
    assert "gemini_disabled_models" in loaded


def test_regression_enrich_uses_cached_entry_without_external_calls(monkeypatch):
    logger = logging.getLogger("test")
    cache = {
        "contacts": {
            "https://maps.google.com/place": {
                "official_website": "https://kies.de",
                "emails_found": "cached@kies.de",
                "phones_found": "",
                "contact_sources": "https://kies.de/kontakt",
                "email_target": "cached@kies.de",
                "email_subject": "x",
                "email_body": "y",
                "email_status": "not_sent",
            }
        },
        "gemini_disabled_models": {},
    }
    row = {"url": "https://maps.google.com/place", "www": "", "nazwa": "Kies"}

    def fail(*args, **kwargs):
        raise AssertionError("External lookup should not be called when cache hit exists")

    monkeypatch.setattr(scraper, "search_official_website_with_serper", fail)
    monkeypatch.setattr(scraper, "collect_contacts_from_website", fail)

    out = scraper.enrich_row_with_contacts(row, cache, logger)
    assert out["emails_found"] == "cached@kies.de"


def test_regression_enrich_marks_maps_rejected_when_identical_contacts(monkeypatch):
    logger = logging.getLogger("test")
    cache = {"contacts": {}, "serper": {}}
    row = {
        "url": "https://maps.google.com/place2",
        "www": "https://kbi.de",
        "telefon": "+49 7229 60101",
        "nazwa": "KBI Kieswerk",
    }

    monkeypatch.setattr(
        scraper,
        "collect_contacts_from_website",
        lambda website, _logger: {
            "emails": ["info@kbi.de"],
            "phones": ["+49 7229 60101"],
            "website": "https://kbi.de",
            "source_urls": ["https://kbi.de/kontakt"],
        },
    )

    out = scraper.enrich_row_with_contacts(row, cache, logger)
    assert out["maps_contact_rejected"] == "yes"
    assert out["contact_source"] == "serper_bs4"


def test_regression_serper_cache_accepts_dict_shape(monkeypatch):
    logger = logging.getLogger("test")
    cache = {"serper": {"Kies Nord": {"url": "https://kiesnord.de", "score": 77}}}
    got = scraper.search_official_website_with_serper("Kies Nord", "", logger, cache)
    assert got == "https://kiesnord.de"
