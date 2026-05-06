"""
pytest 設定檔：收集所有 .bas 檔案路徑供測試使用。
"""
import pytest
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent
MODULES_DIR = REPO_ROOT / "模組"


def collect_bas_files():
    return sorted(MODULES_DIR.rglob("*.bas"))


def pytest_configure(config):
    config.addinivalue_line("markers", "excel: 需要桌面版 Excel 才能執行的測試")


@pytest.fixture(params=collect_bas_files(), ids=lambda p: p.relative_to(REPO_ROOT).as_posix())
def bas_file(request):
    """提供所有 .bas 檔案路徑，供 test_static.py 參數化使用。"""
    return request.param
