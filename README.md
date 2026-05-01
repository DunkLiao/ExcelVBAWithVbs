# ExcelVBAWithVbs

使用 VBScript（.vbs）自動化 Microsoft Excel 的範例集，無需開啟 VBA 編輯器，直接以命令提示字元執行即可。

## 環境需求

- Windows 作業系統
- Microsoft Excel（已安裝）
- Windows Script Host（Windows 內建，無需額外安裝）

## 執行方式

在命令提示字元中切換至專案目錄後執行：

```cmd
cscript Excel\<檔案名稱>.vbs
```

## 範例列表

### Excel

| 檔案                                           | 說明                                                 |
| ---------------------------------------------- | ---------------------------------------------------- |
| [CreateBarChart.vbs](Excel/CreateBarChart.vbs) | 自動建立 2025 年各月銷售額長條圖，並將結果儲存至桌面 |

## 範例說明

### CreateBarChart.vbs

自動在 Excel 中：

1. 建立新活頁簿並填入 12 個月的銷售額示範資料
2. 插入**群組直條圖**，設定圖表標題、X/Y 軸標籤與資料標籤
3. 將結果另存為 `BarChartExample.xlsx` 至桌面

```cmd
cscript Excel\CreateBarChart.vbs
```

## 授權

MIT License
