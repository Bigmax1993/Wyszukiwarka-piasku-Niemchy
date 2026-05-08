import logging

import germany_sand_gravel_scraper as scraper


def test_parse_card_text_splits_fields():
    category, address, status = scraper.parse_card_text(
        "Kieswerk · Musterstrasse 1 · Voruebergehend geschlossen"
    )
    assert category == "Kieswerk"
    assert address == "Musterstrasse 1"
    assert status == "Voruebergehend geschlossen"


def test_extract_open_status_detects_german_closed():
    got = scraper.extract_open_status("Heute: Voruebergehend geschlossen")
    assert got == "Vorübergehend geschlossen"


def test_find_emails_filters_image_suffix():
    txt = "Kontakt: info@example.de logo@site.png sales@example.de."
    got = scraper.find_emails_in_text(txt)
    assert "info@example.de" in got
    assert "sales@example.de" in got
    assert "logo@site.png" not in got


def test_find_phones_requires_enough_digits():
    txt = "Tel +49 152 236 55399, Fax 12-34"
    got = scraper.find_phones_in_text(txt)
    assert any("+49 152 236 55399" in p for p in got)
    assert all("12-34" not in p for p in got)


def test_normalize_website_adds_scheme():
    assert scraper.normalize_website("example.de") == "https://example.de"
    assert scraper.normalize_website("https://example.de") == "https://example.de"


def test_search_url_contains_term_and_coordinates():
    url = scraper.search_url("kieswerk", 50.1, 8.7)
    assert "google.com/maps/search/" in url
    assert "kieswerk+deutschland" in url
    assert "@50.1,8.7" in url


def test_frange_includes_stop_when_exact():
    got = list(scraper.frange(1.0, 1.2, 0.1))
    assert got == [1.0, 1.1]


def test_collect_contacts_from_website_empty_input():
    logger = logging.getLogger("test")
    got = scraper.collect_contacts_from_website("", logger)
    assert got["emails"] == []
    assert got["phones"] == []
    assert got["website"] == ""


def test_score_serper_candidate_penalizes_bad_domains():
    score_bad = scraper.score_serper_candidate(
        "https://facebook.com/company", "Company", "", "Kies Nord"
    )
    score_good = scraper.score_serper_candidate(
        "https://kiesnord.de/kontakt", "Kies Nord", "Impressum", "Kies Nord"
    )
    assert score_good > score_bad


def test_compute_contact_quality_score_prefers_complete_contact():
    low = scraper.compute_contact_quality_score(
        {"email_target": "", "phones_found": "", "official_website": ""}
    )
    high = scraper.compute_contact_quality_score(
        {
            "email_target": "kontakt@kies.de",
            "phones_found": "+49 30 123456",
            "official_website": "https://kies.de",
            "full_address": "Berlin",
            "serper_source_score": 40,
        }
    )
    assert high > low


def test_email_local_part_and_domain_helpers():
    assert scraper.get_email_domain("sales@kies.de") == "kies.de"
    assert scraper.get_email_local_part("sales@kies.de") == "sales"
    assert scraper.is_email_role_based_or_system("noreply@kies.de") is True


def test_reconcile_contact_sources_prefers_website_when_same_contact():
    row = {"telefon": "+49 7229 60101", "www": "https://kbi.example.de"}
    collected = {
        "website": "https://kbi.example.de",
        "phones": ["+49 7229 60101"],
        "emails": [],
        "source_urls": [],
    }
    out = scraper.reconcile_contact_sources(row, collected)
    assert out["contact_source"] == "serper_bs4"
    assert out["maps_contact_rejected"] == "yes"
