# VBA 儲存庫

本儲存庫專門存放 Microsoft Office VBA 程式碼，包含 Excel、Word、Outlook 自動化範例與工具。

# CRITICAL RULES FOR AI AGENTS

## VBA 語法正確性規則

所有 VBA 程式碼必須符合「可直接匯入 VBA 編輯器並通過編譯」的標準。

### 必做規則

- 每個模組第一行必須使用 `Option Explicit`
- 所有變數必須明確宣告
- 不得使用未宣告常數，除非該常數為 Excel VBA 內建常數
- 不得產生偽 VBA、類 Python 語法、類 JavaScript 語法
- 不得只追求示意，必須產出可執行版本
- 產生完整 Sub / Function，不得只產生片段
- 所有括號、引號、續行符號 `_` 必須完整正確
- 每個 `If` 必須有對應 `End If`
- 每個 `For` 必須有對應 `Next`
- 每個 `With` 必須有對應 `End With`
- 每個 `Sub` / `Function` 必須有對應 `End Sub` / `End Function`

---

## Excel VBA 相容性規則

產生 Excel VBA 時，必須符合 Excel 物件模型。

---

## VBA 編譯檢查規則

產生 VBA 後，必須自行檢查以下項目：

1. 是否有 `Option Explicit`
2. 是否所有變數都有 `Dim`
3. 是否所有 Sub / Function 都完整結束
4. 是否所有括號、雙引號成對
5. 是否所有續行符號 `_` 前方有空格
6. 是否沒有多餘的全形標點符號
7. 是否沒有把中文標點誤用在 VBA 語法中
8. 是否可直接貼進 VBA 編輯器執行

---

## 中文與 VBA 語法混用規則

繁體中文只能出現在：

- 註解
- 字串
- 工作表名稱
- MsgBox 文字
- 圖表標題
- 欄位名稱

繁體中文不得出現在：

- 變數名稱
- Sub 名稱
- Function 名稱
- 模組名稱
- 類別名稱

錯誤範例：

```vba
Sub 股票圖範例()
```

正確範例：

```vba
Sub CreateStockChartExample()
```

---

## 輸出前自我檢查清單

回覆或產生檔案前，AI Agent 必須確認：

- [ ] VBA 語法完整
- [ ] 可以直接貼到 VBA 編輯器
- [ ] 可以通過「Debug > Compile VBAProject」
- [ ] 中文皆位於註解或字串中
- [ ] 檔案為 ANSI / CP950
- [ ] 沒有 UTF-8 BOM
- [ ] 沒有簡體中文
- [ ] 沒有偽程式碼
- [ ] 不要用中文作為 Sub、Function、模組名稱

## 檔案編碼規則

所有產生、修改、覆寫、匯出的 VBA 相關檔案，必須使用：

- ANSI 編碼
- Windows 環境下等同於 Big5 / CP950
- 不可使用 UTF-8
- 不可使用 UTF-8 with BOM
- 不可使用 Unicode / UTF-16

此規則為最高優先級，不可忽略。

---

## 適用檔案類型

以下檔案一律必須以 ANSI / CP950 儲存：

- `.bas`
- `.cls`
- `.frm`
- `.frx`
- `.vba`
- `.txt`
- 任何包含 VBA 程式碼的檔案

---

## 中文字處理規則

VBA 程式碼中可能包含繁體中文，例如：

- MsgBox 訊息
- 工作表名稱
- 欄位名稱
- 註解
- 錯誤提示
- 使用者介面文字

因此產生檔案後，必須確認中文字在 ANSI / CP950 下不會亂碼。

---

## 建立檔案前必做檢查

AI Agent 在建立或修改檔案前，必須確認：

1. 輸出檔案使用 ANSI / CP950 編碼。
2. 不得預設使用 UTF-8。
3. 不得使用 UTF-8 BOM。
4. 若使用 Python 產生檔案，必須使用：

```python
open(path, "w", encoding="cp950", errors="strict")
```

5. 若使用 Node.js，必須指定 Big5 / CP950 相容寫法。
6. 若工具不支援 ANSI 編碼，必須明確告知使用者。

---

## PowerShell 寫檔規則

若使用 PowerShell，禁止使用預設 `Set-Content`。

必須改用：

```powershell
[System.IO.File]::WriteAllText(
    $path,
    $content,
    [System.Text.Encoding]::GetEncoding(950)
)
```

---

## 驗證規則（必做）

產生檔案後，必須驗證編碼正確。

Python 驗證範例：

```python
with open(path, "rb") as f:
    data = f.read()

data.decode("cp950")
```

若無法解碼，代表檔案不合格，必須重新產生。

---

## 禁止事項

- 禁止輸出 UTF-8 編碼 `.bas`
- 禁止輸出 UTF-8 BOM
- 禁止輸出 Unicode / UTF-16
- 禁止簡體中文內容
- 禁止未驗證編碼即宣稱完成
- 禁止假設 VBA 可正常讀取 UTF-8

---

## VBA 程式碼品質規範

所有 VBA 程式碼必須符合：

- 使用 `Option Explicit`
- 所有變數明確宣告
- 具備錯誤處理
- 使用繁體中文註解
- 模組名稱清楚
- 可直接匯入 VBA 編輯器

---

## 標準輸出格式

當使用者要求 VBA 程式碼時，優先輸出：

1. `.bas` 匯入檔
2. 完整可執行 VBA 程式碼
3. 已確認 ANSI / CP950 編碼版本
4. 含使用說明

---

# 專案說明

本儲存庫提供 VBA 範例與工具，協助使用者自動化 Microsoft Office 應用程式。

適合：

- Excel 報表自動化
- Word 文件產生
- Outlook 郵件處理
- 批次資料整理
- 企業內部流程工具開發
