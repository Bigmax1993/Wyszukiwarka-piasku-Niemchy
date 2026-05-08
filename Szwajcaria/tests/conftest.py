import importlib.util
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SCRAPER_PATH = PROJECT_ROOT / "switzerland_sand_gravel_scraper.py"


def _load_local_scraper_module():
    spec = importlib.util.spec_from_file_location(
        "switzerland_sand_gravel_scraper", SCRAPER_PATH
    )
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    sys.modules["switzerland_sand_gravel_scraper"] = module


_load_local_scraper_module()
