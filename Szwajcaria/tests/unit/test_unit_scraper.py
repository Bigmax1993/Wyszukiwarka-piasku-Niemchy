import logging

import switzerland_sand_gravel_scraper as scraper


def test_search_url_contains_schweiz():
    url = scraper.search_url("kieswerk", 46.9, 7.4)
    assert "kieswerk+schweiz" in url


def test_country_restriction_is_ch():
    assert scraper.COUNTRY_RESTRICTION == "CH"
    assert scraper.SERPER_COUNTRY == "ch"


def test_collect_contacts_from_website_empty_input():
    logger = logging.getLogger("test")
    got = scraper.collect_contacts_from_website("", logger)
    assert got["emails"] == []
    assert got["phones"] == []
    assert got["website"] == ""
