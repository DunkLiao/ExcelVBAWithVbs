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
cscript charts\<檔案名稱>.vbs
```

## 範例列表

### Excel

| 檔案                                           | 說明                                                 |
| ---------------------------------------------- | ---------------------------------------------------- |
| [CreateBarChart.vbs](Excel/CreateBarChart.vbs) | 自動建立 2025 年各月銷售額長條圖，並將結果儲存至桌面 |

### Charts（圖表範例集）

| 檔案                                                                             | 圖表類型     | 示範主題               |
| -------------------------------------------------------------------------------- | ------------ | ---------------------- |
| [CreateColumnChart.vbs](charts/CreateColumnChart.vbs)                           | 群組直條圖   | 三個產品的季度業績比較 |
| [CreateBarChart.vbs](charts/CreateBarChart.vbs)                                 | 群組橫條圖   | 各部門員工人數         |
| [CreateLineChart.vbs](charts/CreateLineChart.vbs)                               | 折線圖       | 全年各月平均氣溫       |
| [CreatePieChart.vbs](charts/CreatePieChart.vbs)                                 | 圓餅圖       | 智慧型手機市場佔有率   |
| [CreateAreaChart.vbs](charts/CreateAreaChart.vbs)                               | 區域圖       | 網站月流量趨勢         |
| [CreateScatterChart.vbs](charts/CreateScatterChart.vbs)                         | 散佈圖（XY）| 身高與體重分布         |
| [CreateDoughnutChart.vbs](charts/CreateDoughnutChart.vbs)                       | 環圈圖       | 年度預算分配           |
| [CreateRadarChart.vbs](charts/CreateRadarChart.vbs)                             | 雷達圖       | 員工技能評分比較       |
| [CreateStackedColumnChart.vbs](charts/CreateStackedColumnChart.vbs)             | 堆疊直條圖         | 季度收支堆疊分析         |
| [CreatePieExplodedChart.vbs](charts/CreatePieExplodedChart.vbs)                 | 分裂圓餅圖         | 月度生活費用類別佔比     |
| [CreateStackedBarChart.vbs](charts/CreateStackedBarChart.vbs)                   | 堆疊橫條圖         | 各地區產品銷售堆疊       |
| [CreateStackedAreaChart.vbs](charts/CreateStackedAreaChart.vbs)                 | 堆疊區域圖         | 月度電力來源組成         |
| [CreateBubbleChart.vbs](charts/CreateBubbleChart.vbs)                           | 泡泡圖             | 城市 GDP、人口與幸福指數 |
| [CreateColumn3DChart.vbs](charts/CreateColumn3DChart.vbs)                       | 3D 群組直條圖      | 季度線上 vs 實體銷售比較 |
| [CreatePie3DChart.vbs](charts/CreatePie3DChart.vbs)                             | 3D 圓餅圖          | 公司資源分配比例         |
| [CreateLineMarkersChart.vbs](charts/CreateLineMarkersChart.vbs)                 | 含資料點折線圖     | 雙城市全年氣溫對比       |
| [CreateSurfaceChart.vbs](charts/CreateSurfaceChart.vbs)                         | 曲面圖（3D）       | 溫度與壓力下的反應速率   |
| [CreateBarStacked100Chart.vbs](charts/CreateBarStacked100Chart.vbs)             | 百分比堆疊橫條圖   | 各部門費用結構比例       |
| [CreateColumnStacked100Chart.vbs](charts/CreateColumnStacked100Chart.vbs)       | 百分比堆疊直條圖   | 各季品牌市佔率變化       |
| [CreateScatterSmoothChart.vbs](charts/CreateScatterSmoothChart.vbs)             | 平滑線散佈圖       | 產品使用時數與效能衰減   |

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
