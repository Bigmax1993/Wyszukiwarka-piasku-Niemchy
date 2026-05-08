import switzerland_sand_gravel_scraper as scraper


def test_gemini_model_config_is_present():
    assert isinstance(scraper.GEMINI_MODEL, str)
    assert scraper.GEMINI_MODEL != ""


def test_discovery_terms_include_schweiz():
    assert any("schweiz" in x.lower() for x in scraper.SERPER_DISCOVERY_TERMS)
