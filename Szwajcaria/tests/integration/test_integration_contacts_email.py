import logging

import switzerland_sand_gravel_scraper as scraper


def test_generate_email_fallback_contains_signature(monkeypatch):
    logger = logging.getLogger("test")
    monkeypatch.delenv("GOOGLE_AI_STUDIO_API_KEY", raising=False)
    subject, body = scraper.generate_email_content_gemini("Kies Schweiz", logger, cache={})
    assert "Sand" in subject
    assert scraper.EMAIL_SIGNATURE in body
