"""
Phase 1：靜態文字層級檢查。
不需要 Excel，直接分析 .bas 檔案的 bytes 與文字內容。

執行方式：
    pytest tests/test_static.py -v
"""
import re
import pytest
from pathlib import Path


# ---------------------------------------------------------------------------
# 輔助函數
# ---------------------------------------------------------------------------

def read_bytes(bas_file: Path) -> bytes:
    return bas_file.read_bytes()


def decode_cp950(data: bytes) -> list[str]:
    """將 bytes 以 CP950 解碼後，回傳行列表（去掉行尾換行）。"""
    return data.decode("cp950").splitlines()


# 簡體中文 Unicode 範圍（CJK Unified Ideographs 中屬於簡體常用字）
# 使用一組常見簡體專用字的 regex 作為啟發式偵測
# 完整判斷需要字典，這裡用「GB2312 有但 Big5 無法對應」的字元範圍近似
_SIMPLIFIED_ONLY_PATTERN = re.compile(
    r"[\u4e2a\u4e2a\u4f53\u5c42\u5c71\u5c8c\u5e76\u5f00\u5f52\u5f62\u5f84"
    r"\u6240\u62a5\u6309\u6362\u6765\u6784\u6797\u6c49\u6d4b\u6d4e\u6d88"
    r"\u7ec4\u7ed3\u7edf\u7f16\u8303\u8868\u88c5\u89e3\u8fde\u8ff0\u9009"
    r"\u9014\u914d\u9274\u9489\u9500\u9636\u964d\u97f3\u9881\u9898\u9a6c"
    r"\u9f99\u94fe\u9489]"
)

# 中文字元（CJK 統一漢字）用於偵測 Sub/Function 名稱是否含中文
_CJK_PATTERN = re.compile(r"[\u4e00-\u9fff\u3400-\u4dbf]")


# ---------------------------------------------------------------------------
# 測試：CP950 編碼
# ---------------------------------------------------------------------------

def test_encoding_is_cp950(bas_file):
    """檔案必須能以 CP950 解碼，確認不是 UTF-8 或其他編碼。"""
    data = read_bytes(bas_file)
    try:
        data.decode("cp950")
    except UnicodeDecodeError as e:
        pytest.fail(f"無法以 CP950 解碼：{e}")


# ---------------------------------------------------------------------------
# 測試：無 UTF-8 BOM
# ---------------------------------------------------------------------------

def test_no_utf8_bom(bas_file):
    """檔案前三個 bytes 不得為 UTF-8 BOM（\\xef\\xbb\\xbf）。"""
    data = read_bytes(bas_file)
    assert not data.startswith(b"\xef\xbb\xbf"), "檔案含有 UTF-8 BOM，請以 ANSI/CP950 儲存"


# ---------------------------------------------------------------------------
# 測試：Option Explicit
# ---------------------------------------------------------------------------

def test_option_explicit(bas_file):
    """
    模組必須包含 'Option Explicit'。
    允許 .bas 匯出格式在開頭有 'Attribute VB_*' 標頭行，
    但 Option Explicit 必須出現在 Attribute 區塊結束後的第一個有效行。
    """
    data = read_bytes(bas_file)
    lines = decode_cp950(data)
    # 略過 Attribute VB_* 標頭行（VBA 匯出格式的合法開頭）
    non_attr_lines = [
        l.strip() for l in lines
        if l.strip() and not l.strip().startswith("Attribute VB_")
    ]
    first_code_line = non_attr_lines[0] if non_attr_lines else None
    assert first_code_line == "Option Explicit", (
        f"第一個程式碼行應為 'Option Explicit'，實際為：{first_code_line!r}"
    )


# ---------------------------------------------------------------------------
# 輔助：計算關鍵字出現次數（忽略註解行與字串）
# ---------------------------------------------------------------------------

def _count_keyword_lines(lines: list[str], pattern: re.Pattern) -> int:
    count = 0
    for line in lines:
        stripped = line.strip()
        # 忽略純註解行
        if stripped.startswith("'"):
            continue
        # 移除行內註解（簡單處理：取第一個 ' 前的部分，但需排除字串內的 '）
        code_part = stripped.split("'")[0]
        if pattern.search(code_part):
            count += 1
    return count


_SUB_OPEN = re.compile(r"^(Public\s+|Private\s+|Friend\s+)?Sub\s+\w", re.IGNORECASE)
_SUB_CLOSE = re.compile(r"^End\s+Sub\b", re.IGNORECASE)
_FUNC_OPEN = re.compile(r"^(Public\s+|Private\s+|Friend\s+)?Function\s+\w", re.IGNORECASE)
_FUNC_CLOSE = re.compile(r"^End\s+Function\b", re.IGNORECASE)
_WITH_OPEN = re.compile(r"^With\s+", re.IGNORECASE)
_WITH_CLOSE = re.compile(r"^End\s+With\b", re.IGNORECASE)
# If...Then 單行（不需要 End If）
_IF_SINGLE = re.compile(r"^If\b.+\bThen\b.+$", re.IGNORECASE)
_IF_OPEN = re.compile(r"^If\b.+\bThen\s*$", re.IGNORECASE)
_IF_CLOSE = re.compile(r"^End\s+If\b", re.IGNORECASE)


# ---------------------------------------------------------------------------
# 測試：Sub / End Sub 成對
# ---------------------------------------------------------------------------

def test_sub_end_sub_paired(bas_file):
    """Sub 與 End Sub 必須數量相同。"""
    data = read_bytes(bas_file)
    lines = decode_cp950(data)
    opens = _count_keyword_lines(lines, _SUB_OPEN)
    closes = _count_keyword_lines(lines, _SUB_CLOSE)
    assert opens == closes, f"Sub 有 {opens} 個，End Sub 有 {closes} 個，數量不符"


# ---------------------------------------------------------------------------
# 測試：Function / End Function 成對
# ---------------------------------------------------------------------------

def test_function_end_function_paired(bas_file):
    """Function 與 End Function 必須數量相同。"""
    data = read_bytes(bas_file)
    lines = decode_cp950(data)
    opens = _count_keyword_lines(lines, _FUNC_OPEN)
    closes = _count_keyword_lines(lines, _FUNC_CLOSE)
    assert opens == closes, f"Function 有 {opens} 個，End Function 有 {closes} 個，數量不符"


# ---------------------------------------------------------------------------
# 測試：With / End With 成對
# ---------------------------------------------------------------------------

def test_with_end_with_paired(bas_file):
    """With 與 End With 必須數量相同。"""
    data = read_bytes(bas_file)
    lines = decode_cp950(data)
    opens = _count_keyword_lines(lines, _WITH_OPEN)
    closes = _count_keyword_lines(lines, _WITH_CLOSE)
    assert opens == closes, f"With 有 {opens} 個，End With 有 {closes} 個，數量不符"


# ---------------------------------------------------------------------------
# 輔助：合併 VBA 續行符號 _ 的邏輯行
# ---------------------------------------------------------------------------

def _join_continuation_lines(lines: list[str]) -> list[str]:
    """將以 _ 結尾的行與下一行合併，還原成完整的邏輯行後再做分析。"""
    result = []
    i = 0
    while i < len(lines):
        s = lines[i]
        stripped = s.strip()
        if stripped.startswith("'"):
            result.append(s)
            i += 1
            continue
        code = stripped.split("'")[0]
        while code.rstrip().endswith("_"):
            code = code.rstrip()[:-1].rstrip()
            i += 1
            if i < len(lines):
                next_stripped = lines[i].strip()
                if next_stripped.startswith("'"):
                    break
                code = code + " " + next_stripped.split("'")[0]
            else:
                break
        result.append(code)
        i += 1
    return result


# ---------------------------------------------------------------------------
# 測試：If / End If 成對（排除單行 If）
# ---------------------------------------------------------------------------

def test_if_end_if_paired(bas_file):
    """多行 If...Then 與 End If 必須數量相同（單行 If...Then...Else 不計）。"""
    data = read_bytes(bas_file)
    raw_lines = decode_cp950(data)
    lines = _join_continuation_lines(raw_lines)
    opens = 0
    closes = 0
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("'"):
            continue
        code_part = stripped.split("'")[0].strip()
        if _IF_CLOSE.match(code_part):
            closes += 1
        elif _IF_OPEN.match(code_part) and not _IF_SINGLE.match(code_part):
            opens += 1
    assert opens == closes, f"多行 If 有 {opens} 個，End If 有 {closes} 個，數量不符"


# ---------------------------------------------------------------------------
# 測試：續行符號 _ 前方必須有空格
# ---------------------------------------------------------------------------

_BAD_CONTINUATION = re.compile(r"[^ \t]_\s*$")


def test_line_continuation_has_space(bas_file):
    """續行符號 _ 前方必須有空格（' _' 而非 '_'）。"""
    data = read_bytes(bas_file)
    lines = decode_cp950(data)
    bad_lines = []
    for i, line in enumerate(lines, start=1):
        if line.strip().startswith("'"):
            continue
        if _BAD_CONTINUATION.search(line):
            bad_lines.append((i, line.rstrip()))
    assert not bad_lines, (
        "以下行的續行符號 _ 前方缺少空格：\n"
        + "\n".join(f"  行 {n}: {l}" for n, l in bad_lines)
    )


# ---------------------------------------------------------------------------
# 測試：Sub / Function 名稱不得含中文
# ---------------------------------------------------------------------------

_SUB_NAME = re.compile(
    r"^(?:Public\s+|Private\s+|Friend\s+)?Sub\s+(\S+)\s*\(", re.IGNORECASE
)
_FUNC_NAME = re.compile(
    r"^(?:Public\s+|Private\s+|Friend\s+)?Function\s+(\S+)\s*\(", re.IGNORECASE
)


def test_no_chinese_sub_function_names(bas_file):
    """Sub 名稱與 Function 名稱不得含繁體或簡體中文字元。"""
    data = read_bytes(bas_file)
    lines = decode_cp950(data)
    bad = []
    for i, line in enumerate(lines, start=1):
        stripped = line.strip()
        if stripped.startswith("'"):
            continue
        for pat in (_SUB_NAME, _FUNC_NAME):
            m = pat.match(stripped)
            if m and _CJK_PATTERN.search(m.group(1)):
                bad.append((i, stripped))
    assert not bad, (
        "以下行使用中文作為 Sub/Function 名稱（應改為英文）：\n"
        + "\n".join(f"  行 {n}: {l}" for n, l in bad)
    )
