import germany_sand_gravel_scraper as scraper


def test_domain_daily_limit_helpers():
    cache = {"email_domain_daily": {}}
    today, sent, remaining = scraper.get_domain_remaining_daily_limit(cache, "kies.de")
    assert sent == 0
    assert remaining == scraper.EMAIL_PER_DOMAIN_DAILY_LIMIT

    scraper.increase_domain_daily_counter(cache, "kies.de", 1)
    _, sent2, remaining2 = scraper.get_domain_remaining_daily_limit(cache, "kies.de")
    assert sent2 == 1
    assert remaining2 == scraper.EMAIL_PER_DOMAIN_DAILY_LIMIT - 1
    assert today in cache["email_domain_daily"]


def test_suppression_helpers():
    cache = {"email_suppression": {}}
    target = "noreply@kies.de"
    assert scraper.is_email_role_based_or_system(target) is True
    assert scraper.is_suppressed_target(cache, target) is False
    scraper.mark_suppressed_target(cache, target, "test")
    assert scraper.is_suppressed_target(cache, target) is True


def test_choose_variants_are_stable():
    name = "Kies Nord GmbH"
    s1 = scraper.choose_subject_variant(name)
    s2 = scraper.choose_subject_variant(name)
    p1 = scraper.choose_prompt_variant(name)
    p2 = scraper.choose_prompt_variant(name)
    assert s1 == s2
    assert p1 == p2


def test_serper_limit_flag_helpers():
    cache = {"serper_daily": {}, "serper_limit_reached": {}}
    assert scraper.is_serper_limit_reached_today(cache) is False
    scraper.mark_serper_limit_reached_today(cache)
    assert scraper.is_serper_limit_reached_today(cache) is True
