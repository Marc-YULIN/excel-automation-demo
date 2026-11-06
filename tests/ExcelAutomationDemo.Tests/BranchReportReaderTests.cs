using ClosedXML.Excel;

namespace ExcelAutomationDemo.Tests;

public class BranchReportReaderTests
{
    private const string FileName = "分店月報_台北店_2026-10.xlsx";

    /// <summary>標題在第 3 列，資料含各種人工輸入問題。</summary>
    private static XLWorkbook MessyWorkbook()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("銷售明細");
        ws.Cell("A1").Value = "台北店 2026 年 10 月銷售月報";
        ws.Cell("A3").Value = "日期"; ws.Cell("B3").Value = "品項"; ws.Cell("C3").Value = "數量";
        ws.Cell("D3").Value = "單價"; ws.Cell("E3").Value = "金額";

        // 正常資料
        ws.Cell("A4").Value = new DateTime(2026, 10, 1); ws.Cell("B4").Value = "珍珠奶茶";
        ws.Cell("C4").Value = 10; ws.Cell("E4").Value = 550;
        // 品項前後空白、數量存成文字、日期是文字
        ws.Cell("A5").Value = "2026/10/02"; ws.Cell("B5").Value = "  紅茶拿鐵 ";
        ws.Cell("C5").Value = "3"; ws.Cell("E5").Value = "NT$ 1,800";
        // 空白列
        // 金額是 N/A → 異常
        ws.Cell("A7").Value = new DateTime(2026, 10, 3); ws.Cell("B7").Value = "冬瓜檸檬";
        ws.Cell("C7").Value = 2; ws.Cell("E7").Value = "N/A";
        // 日期亂打、品項空白 → 異常（兩個問題合併成一筆）
        ws.Cell("A8").Value = "上週五"; ws.Cell("C8").Value = 1; ws.Cell("E8").Value = 50;
        // 合計列 → 略過
        ws.Cell("B9").Value = "合計"; ws.Cell("C9").FormulaA1 = "SUM(C4:C8)"; ws.Cell("E9").FormulaA1 = "SUM(E4:E8)";
        return wb;
    }

    [Fact]
    public void ReadsAndCleansValidRows()
    {
        using var wb = MessyWorkbook();
        var result = BranchReportReader.Read(wb, FileName);

        Assert.Equal(2, result.Records.Count);
        Assert.Equal(new SaleRecord("台北店", new DateOnly(2026, 10, 1), "珍珠奶茶", 10, 550), result.Records[0]);
        Assert.Equal(new SaleRecord("台北店", new DateOnly(2026, 10, 2), "紅茶拿鐵", 3, 1800), result.Records[1]);
    }

    [Fact]
    public void ReportsUnparseableRowsAndSkipsTotals()
    {
        using var wb = MessyWorkbook();
        var issues = BranchReportReader.Read(wb, FileName).Issues;

        Assert.Equal(2, issues.Count);
        Assert.Equal(7, issues[0].Row);
        Assert.Equal("金額無法辨識", issues[0].Problem);
        Assert.Equal(8, issues[1].Row);
        Assert.Equal("日期無法辨識、品項空白", issues[1].Problem);
        Assert.Contains("上週五", issues[1].RawContent);
    }

    [Fact]
    public void HandlesDifferentColumnOrder()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("工作表1");
        ws.Cell("A1").Value = "金額"; ws.Cell("B1").Value = "數量";
        ws.Cell("C1").Value = "品項"; ws.Cell("D1").Value = "日期";
        ws.Cell("A2").Value = 120; ws.Cell("B2").Value = 2; ws.Cell("C2").Value = "芒果冰沙";
        ws.Cell("D2").Value = 46296; // 沒套日期格式的 Excel 日期序號 = 2026/10/01

        var result = BranchReportReader.Read(wb, "高雄店.xlsx");

        Assert.Empty(result.Issues);
        Assert.Equal(new SaleRecord("高雄店", new DateOnly(2026, 10, 1), "芒果冰沙", 2, 120), Assert.Single(result.Records));
    }

    [Fact]
    public void ReportsMissingHeader()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1").Cell("A1").Value = "這不是月報";

        var result = BranchReportReader.Read(wb, "其他.xlsx");

        Assert.Empty(result.Records);
        var issue = Assert.Single(result.Issues);
        Assert.Null(issue.Row);
        Assert.StartsWith("找不到標題列", issue.Problem);
    }

    [Fact]
    public void RejectsFractionalQuantity()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("銷售明細");
        ws.Cell("A1").Value = "日期"; ws.Cell("B1").Value = "品項"; ws.Cell("C1").Value = "數量"; ws.Cell("D1").Value = "金額";
        ws.Cell("A2").Value = new DateTime(2026, 10, 1); ws.Cell("B2").Value = "珍珠奶茶";
        ws.Cell("C2").Value = 1.5; ws.Cell("D2").Value = 80;

        var result = BranchReportReader.Read(wb, FileName);

        Assert.Empty(result.Records);
        Assert.Equal("數量無法辨識", Assert.Single(result.Issues).Problem);
    }

    [Theory]
    [InlineData("分店月報_台北店_2026-10.xlsx", "台北店")]
    [InlineData("分店月報_新竹 竹北店_2026-10.xlsx", "新竹 竹北店")]
    [InlineData("台中店十月.xlsx", "台中店十月")]
    public void BranchNameFromFile(string fileName, string expected) =>
        Assert.Equal(expected, BranchReportReader.BranchNameFromFile(fileName));
}
