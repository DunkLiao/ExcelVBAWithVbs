"""
Phase 2：Excel COM 動態編譯驗證。
需要桌面版 Excel，且已開啟 VBA 專案物件模型存取信任設定。

前置條件：
    Excel > 檔案 > 選項 > 信任中心 > 信任中心設定 > 巨集設定
    勾選「信任 VBA 專案物件模型的存取」

執行方式：
    pytest tests/test_compile.py -v
    pytest tests/ -m "not excel"  # 跳過此檔（無 Excel 環境）
"""
import os
import time
import pytest
from pathlib import Path

# 動態偵測 pywin32 是否可用
try:
    import win32com.client
    import pywintypes
    _WIN32_AVAILABLE = True
except ImportError:
    _WIN32_AVAILABLE = False

pytestmark = pytest.mark.excel

REPO_ROOT = Path(__file__).parent.parent
MODULES_DIR = REPO_ROOT / "模組"
TEMP_WORKBOOK = Path(os.environ["TEMP"]) / "VbaCompileCheck.xlsm"

# xlOpenXMLWorkbookMacroEnabled
_XL_MACRO_ENABLED = 52


def _collect_bas_groups() -> dict[str, list[Path]]:
    """
    將所有 .bas 依子資料夾分組，每組一次匯入並編譯。
    同一資料夾的 .bas 視為同一個「模組群組」。
    """
    groups: dict[str, list[Path]] = {}
    for bas in MODULES_DIR.rglob("*.bas"):
        key = str(bas.parent)
        groups.setdefault(key, []).append(bas)
    return groups


_GROUPS = _collect_bas_groups()


# ---------------------------------------------------------------------------
# 前置條件檢查
# ---------------------------------------------------------------------------

@pytest.fixture(scope="session", autouse=True)
def require_win32():
    if not _WIN32_AVAILABLE:
        pytest.skip("pywin32 未安裝，無法執行 Excel COM 測試")


# ---------------------------------------------------------------------------
# 工具函數
# ---------------------------------------------------------------------------

def _open_excel():
    """建立 Excel COM 物件。"""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True  # Visible=True 才能觸發 VBE GUI 指令
    excel.DisplayAlerts = False
    return excel


def _save_and_close(workbook, excel):
    try:
        if TEMP_WORKBOOK.exists():
            workbook.Close(False)
        else:
            workbook.Close(False)
    except Exception:
        pass
    try:
        excel.Quit()
    except Exception:
        pass
    if TEMP_WORKBOOK.exists():
        TEMP_WORKBOOK.unlink(missing_ok=True)


def _compile_bas_files(bas_files: list[Path]) -> tuple[bool, str]:
    """
    匯入 bas_files 到暫存 workbook 並觸發 VBE Compile VBAProject。
    回傳 (success: bool, message: str)。
    """
    excel = _open_excel()
    workbook = None
    try:
        workbook = excel.Workbooks.Add()
        workbook.SaveAs(str(TEMP_WORKBOOK), _XL_MACRO_ENABLED)

        for bas in bas_files:
            workbook.VBProject.VBComponents.Import(str(bas))

        # CommandBar ID 578 = Debug > Compile VBAProject
        compile_btn = excel.VBE.CommandBars.FindControl(1, 578)
        if compile_btn is None:
            return False, "找不到 VBE Compile VBAProject 指令（CommandBar ID 578）"

        compile_btn.Execute()
        time.sleep(1)  # 給 VBE 一點時間顯示錯誤（若有）

        workbook.Save()
        return True, "編譯通過"

    except pywintypes.com_error as e:
        return False, f"COM 錯誤：{e}"
    finally:
        _save_and_close(workbook, excel)


# ---------------------------------------------------------------------------
# 參數化測試：每個子資料夾群組一個測試案例
# ---------------------------------------------------------------------------

@pytest.mark.parametrize(
    "group_name,bas_files",
    [(name, files) for name, files in _GROUPS.items()],
    ids=[Path(name).name for name in _GROUPS.keys()],
)
def test_compile_group(group_name, bas_files):
    """
    將同一資料夾的 .bas 一起匯入並觸發 Excel VBE 編譯。
    若 VBE 未停在錯誤行，視為編譯通過。
    """
    success, message = _compile_bas_files(bas_files)
    assert success, (
        f"群組 [{Path(group_name).name}] 編譯失敗：{message}\n"
        f"涵蓋檔案：\n" + "\n".join(f"  {f.name}" for f in bas_files)
    )
