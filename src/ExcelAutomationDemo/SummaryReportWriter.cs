using ClosedXML.Excel;

namespace ExcelAutomationDemo;

/// <summary>輸出彙總報表：總覽、品項排行、合併明細、異常資料。</summary>
public static class SummaryReportWriter
{
    private const string MoneyFormat = "\"NT$\"#,##0";
    private const string CountFormat = "#,##0";
    private const string PercentFormat = "0.0%";
    private const string DateFormat = "yyyy/mm/dd";
    private static readonly XLColor HeaderColor = XLColor.FromHtml("#305496");

    public static void Write(
        string path, IReadOnlyList<SaleRecord> records, IReadOnlyList<DataIssue> issues,
        int sourceFileCount, DateTime generatedAt)
    {
        var branches = ReportCalculator.SummarizeBranches(records);
        var ranking = ReportCalculator.RankItems(records);

        using var wb = new XLWorkbook();
        WriteOverview(wb.AddWorksheet("總覽"), records, branches, issues.Count, sourceFileCount, generatedAt);
        WriteRanking(wb.AddWorksheet("品項排行"), ranking, branches.Select(b => b.Branch).ToList());
        WriteDetails(wb.AddWorksheet("合併明細"), records);
        if (issues.Count > 0)
            WriteIssues(wb.AddWorksheet("異常資料"), issues);
        wb.SaveAs(path);
    }

    private static void WriteOverview(
        IXLWorksheet ws, IReadOnlyList<SaleRecord> records, IReadOnlyList<BranchSummary> branches,
        int issueCount, int sourceFileCount, DateTime generatedAt)
    {
        ws.Cell("A1").Value = "分店銷售彙總報表";
        ws.Range("A1:F1").Merge().Style.Font.SetBold().Font.SetFontSize(16);

        var period = records.Count == 0
            ? "無資料"
            : $"{records.Min(r => r.Date):yyyy/MM/dd} – {records.Max(r => r.Date):yyyy/MM/dd}";
        ws.Cell("A2").Value = $"資料期間：{period}　來源檔案：{sourceFileCount} 個　產出時間：{generatedAt:yyyy/MM/dd HH:mm}";
        ws.Range("A2:F2").Merge().Style.Font.SetFontColor(XLColor.Gray);

        const int headerRow = 4;
        WriteHeader(ws, headerRow, ["分店", "交易筆數", "總數量", "總金額", "佔比", "最暢銷品項"]);

        var row = headerRow + 1;
        foreach (var b in branches)
        {
            ws.Cell(row, 1).Value = b.Branch;
            ws.Cell(row, 2).Value = b.Transactions;
            ws.Cell(row, 3).Value = b.Quantity;
            ws.Cell(row, 4).Value = b.Amount;
            ws.Cell(row, 5).Value = b.Share;
            ws.Cell(row, 6).Value = b.TopItem;
            row++;
        }

        // 合計列用公式，業主在 Excel 裡可以自己驗算
        var first = headerRow + 1;
        var last = row - 1;
        ws.Cell(row, 1).Value = "合計";
        foreach (var c in new[] { "B", "C", "D" })
            ws.Cell(row, c).FormulaA1 = $"SUM({c}{first}:{c}{last})";
        ws.Cell(row, 5).Value = branches.Count > 0 ? 1 : 0;
        var totalRow = ws.Range(row, 1, row, 6);
        totalRow.Style.Font.Bold = true;
        totalRow.Style.Border.TopBorder = XLBorderStyleValues.Double;

        ws.Range(first, 2, row, 3).Style.NumberFormat.Format = CountFormat;
        ws.Range(first, 4, row, 4).Style.NumberFormat.Format = MoneyFormat;
        ws.Range(first, 5, row, 5).Style.NumberFormat.Format = PercentFormat;

        if (issueCount > 0)
        {
            ws.Cell(row + 2, 1).Value = $"⚠ 有 {issueCount} 筆資料無法自動判斷，已排除在統計之外，請見「異常資料」工作表。";
            ws.Range(row + 2, 1, row + 2, 6).Merge().Style.Font.SetFontColor(XLColor.FromHtml("#C00000"));
        }

        ExcelLayout.FitColumns(ws, 6, minWidth: 12);
    }

    private static void WriteRanking(IXLWorksheet ws, IReadOnlyList<ItemRanking> ranking, IReadOnlyList<string> branches)
    {
        string[] fixedHeaders = ["排名", "品項", "總數量", "總金額", "佔比"];
        var headers = fixedHeaders.Concat(branches.Select(b => $"{b}金額")).ToArray();
        WriteHeader(ws, 1, headers);

        for (var i = 0; i < ranking.Count; i++)
        {
            var item = ranking[i];
            var row = i + 2;
            ws.Cell(row, 1).Value = item.Rank;
            ws.Cell(row, 2).Value = item.Item;
            ws.Cell(row, 3).Value = item.Quantity;
            ws.Cell(row, 4).Value = item.Amount;
            ws.Cell(row, 5).Value = item.Share;
            for (var b = 0; b < branches.Count; b++)
                ws.Cell(row, 6 + b).Value = item.AmountByBranch[branches[b]];
        }

        var lastRow = ranking.Count + 1;
        if (ranking.Count > 0)
        {
            ws.Range(2, 1, lastRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Range(2, 3, lastRow, 3).Style.NumberFormat.Format = CountFormat;
            ws.Range(2, 4, lastRow, 4).Style.NumberFormat.Format = MoneyFormat;
            ws.Range(2, 5, lastRow, 5).Style.NumberFormat.Format = PercentFormat;
            ws.Range(2, 6, lastRow, headers.Length).Style.NumberFormat.Format = MoneyFormat;

            // 前三名標色，總金額加資料橫條，一眼看出差距
            ws.Range(2, 1, Math.Min(lastRow, 4), headers.Length).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF2CC");
            ws.Range(2, 4, lastRow, 4).AddConditionalFormat().DataBar(XLColor.FromHtml("#5B9BD5")).LowestValue().HighestValue();
        }

        ws.SheetView.FreezeRows(1);
        ExcelLayout.FitColumns(ws, headers.Length, minWidth: 8);
    }

    private static void WriteDetails(IXLWorksheet ws, IReadOnlyList<SaleRecord> records)
    {
        string[] headers = ["分店", "日期", "品項", "數量", "金額"];
        for (var c = 0; c < headers.Length; c++) ws.Cell(1, c + 1).Value = headers[c];

        var sorted = records.OrderBy(r => r.Date).ThenBy(r => r.Branch, StringComparer.Ordinal).ToList();
        for (var i = 0; i < sorted.Count; i++)
        {
            var r = sorted[i];
            ws.Cell(i + 2, 1).Value = r.Branch;
            ws.Cell(i + 2, 2).Value = r.Date.ToDateTime(TimeOnly.MinValue);
            ws.Cell(i + 2, 3).Value = r.Item;
            ws.Cell(i + 2, 4).Value = r.Quantity;
            ws.Cell(i + 2, 5).Value = r.Amount;
        }

        // 轉成 Excel 表格：自帶篩選、斑馬紋，業主可以直接樞紐分析
        var table = ws.Range(1, 1, Math.Max(sorted.Count + 1, 2), headers.Length).CreateTable("合併明細");
        table.Theme = XLTableTheme.TableStyleMedium2;
        ws.Column(2).Style.DateFormat.Format = DateFormat;
        ws.Column(4).Style.NumberFormat.Format = CountFormat;
        ws.Column(5).Style.NumberFormat.Format = MoneyFormat;
        ws.SheetView.FreezeRows(1);
        ExcelLayout.FitColumns(ws, headers.Length, minWidth: 12);
    }

    private static void WriteIssues(IXLWorksheet ws, IReadOnlyList<DataIssue> issues)
    {
        WriteHeader(ws, 1, ["來源檔案", "列號", "問題", "原始內容"], XLColor.FromHtml("#C00000"));
        for (var i = 0; i < issues.Count; i++)
        {
            var issue = issues[i];
            ws.Cell(i + 2, 1).Value = issue.File;
            ws.Cell(i + 2, 2).Value = issue.Row is int r ? r : Blank.Value;
            ws.Cell(i + 2, 3).Value = issue.Problem;
            ws.Cell(i + 2, 4).Value = issue.RawContent;
        }
        ws.SheetView.FreezeRows(1);
        ExcelLayout.FitColumns(ws, 4, minWidth: 10);
    }

    private static void WriteHeader(IXLWorksheet ws, int row, string[] headers, XLColor? color = null)
    {
        for (var c = 0; c < headers.Length; c++) ws.Cell(row, c + 1).Value = headers[c];
        var range = ws.Range(row, 1, row, headers.Length);
        range.Style.Font.Bold = true;
        range.Style.Font.FontColor = XLColor.White;
        range.Style.Fill.BackgroundColor = color ?? HeaderColor;
        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    }
}
