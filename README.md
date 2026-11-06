# excel-automation-demo：多份 Excel 一鍵彙總

![.NET 10](https://img.shields.io/badge/.NET-10-512BD4) ![License: MIT](https://img.shields.io/badge/license-MIT-green)

C# / .NET 10 Excel automation demo: merges messy branch reports into one formatted summary workbook in one click.

把各分店手動維護的 Excel 月報，一鍵彙總成格式化的報表：分店小計、品項排行、合併明細，並自動找出有問題的資料。不用寫巨集、不受 Office 版本影響。

## 解決什麼問題

每個月要把 3 間分店的月報合成一份，手動做通常是這樣：

| 步驟 | 手動整理 | 這個工具 |
|---|---|---|
| 開檔、複製貼上各分店資料 | 逐一開檔、注意不要貼到合計列 | 自動讀取資料夾內所有 `.xlsx` |
| 處理格式不一致 | 品項多空白、數字存成文字，要一筆筆修 | 自動清理 |
| 找出錯誤資料 | 靠肉眼，常常漏掉 | 列在「異常資料」工作表，標出檔名與列號 |
| 算分店小計、品項排行 | 樞紐分析表或公式，每月重做一次 | 自動產生 |
| 套格式（千分位、日期、顏色） | 手動調整 | 自動套用 |
| **花費時間** | **約 30–60 分鐘** | **約 1 秒** |

## 執行畫面

<!-- 截圖：docs/images/excel-before.png（分店月報，含格式問題）、docs/images/excel-after.png（彙總報表） -->

```text
已產生 3 份 2026-10 範例月報：
  sample-input\分店月報_台北店_2026-10.xlsx
  sample-input\分店月報_台中店_2026-10.xlsx
  sample-input\分店月報_高雄店_2026-10.xlsx

讀取 3 個檔案，有效資料 338 筆，異常 3 筆
  台中店：NT$144,000（35.7%）
  台北店：NT$140,660（34.8%）
  高雄店：NT$119,010（29.5%）
  ! 分店月報_台中店_2026-10.xlsx 第 97 列：金額無法辨識
  ! 分店月報_台北店_2026-10.xlsx 第 71 列：金額無法辨識
  ! 分店月報_高雄店_2026-10.xlsx 第 87 列：金額無法辨識

完成，耗時 0.5 秒
彙總報表：彙總報表.xlsx
```

產出的 `彙總報表.xlsx`：

| 工作表 | 內容 |
|---|---|
| 總覽 | 各分店交易筆數、總數量、總金額、佔比、最暢銷品項；合計列用 SUM 公式，可在 Excel 裡自行驗算 |
| 品項排行 | 依金額排名，前三名標色、資料橫條，並列出各分店金額 |
| 合併明細 | 所有分店資料合成一張 Excel 表格，可直接篩選或做樞紐分析 |
| 異常資料 | 無法自動判斷的資料：來源檔案、列號、問題、原始內容 |

## 功能重點

- **不依賴固定格式**：自動找標題列（不一定在第 1 列）、欄位順序不同也能讀
- **自動清理常見手打問題**：品項前後空白、數字存成文字、「NT$ 1,200」這類金額、文字格式的日期
- **不默默算錯**：無法判斷的資料（例如金額填「N/A」）會排除並列出來，而不是當成 0
- **略過人工合計列**：避免重複計算
- **範例資料可重現**：用 [Bogus](https://github.com/bchavez/Bogus) 產生假資料，相同亂數種子產生相同內容
- **有測試**：讀檔清理、彙總計算、端對端流程共 16 個測試

## 執行方式

需要 [.NET 10 SDK](https://dotnet.microsoft.com/download)。

```bash
git clone https://github.com/Marc-YULIN/excel-automation-demo.git
cd excel-automation-demo

# 一鍵示範：產生 3 份範例月報 → 彙總
dotnet run --project src/ExcelAutomationDemo

# 只彙總自己的檔案（放在 my-reports 資料夾）
dotnet run --project src/ExcelAutomationDemo -- merge --dir my-reports --output 十月彙總.xlsx

# 執行測試
dotnet test
```

| 指令／參數 | 說明 |
|---|---|
| （不加指令） | 一鍵示範：產生範例月報後立即彙總 |
| `generate` | 只產生範例月報 |
| `merge` | 只彙總資料夾內的月報 |
| `--dir <資料夾>` | 月報所在資料夾，預設 `sample-input` |
| `--output <檔名>` | 彙總報表檔名，預設 `彙總報表.xlsx` |
| `--month <yyyy-MM>` | 範例資料月份，預設上個月 |
| `--seed <數字>` | 亂數種子，預設 42 |

月報需要包含「日期、品項、數量、金額」四個欄位；分店名稱取自檔名 `分店月報_{分店}_{yyyy-MM}.xlsx`，格式不符時直接用檔名。

## 專案結構

```text
src/ExcelAutomationDemo/
  Program.cs                命令列指令與參數
  SampleDataGenerator.cs    用 Bogus 產生範例月報（含刻意放的格式問題）
  BranchReportReader.cs     讀檔、找標題列、清理資料、記錄異常
  ReportCalculator.cs       分店小計、品項排行（純計算）
  SummaryReportWriter.cs    輸出 4 個工作表的彙總報表
  ExcelLayout.cs            中文友善的欄寬計算
tests/ExcelAutomationDemo.Tests/
```

Excel 讀寫使用 [ClosedXML](https://github.com/ClosedXML/ClosedXML)（MIT 授權，商用免費），電腦不需要安裝 Office。

## 可客製成什麼服務

- **多檔彙總**：各分店、各業務、各月份的報表合併成一份
- **固定格式報表**：依公司格式自動產出週報、月報、對帳單
- **資料檢核**：找出重複、缺漏、金額對不上的資料
- **排程自動執行**：每天或每月固定時間產出，並自動寄 Email
- **串接其他來源**：從資料庫、Google Sheets、系統匯出的 CSV 取資料

## 授權

[MIT](LICENSE)
