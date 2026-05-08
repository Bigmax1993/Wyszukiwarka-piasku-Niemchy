import switzerland_sand_gravel_scraper as scraper


def test_domain_daily_limit_helpers():
    cache = {"email_domain_daily": {}}
    _, sent, remaining = scraper.get_domain_remaining_daily_limit(cache, "kies.ch")
    assert sent == 0
    assert remaining == scraper.EMAIL_PER_DOMAIN_DAILY_LIMIT
