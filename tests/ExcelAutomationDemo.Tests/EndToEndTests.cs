using ClosedXML.Excel;

namespace ExcelAutomationDemo.Tests;

/// <summary>產生範例 → 讀回 → 輸出報表，整條流程跑一次。</summary>
public sealed class EndToEndTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "excel-demo-tests-" + Guid.NewGuid().ToString("N"));
    private static readonly DateOnly Month = new(2026, 10, 1);

    public void Dispose()
    {
        if (Directory.Exists(_dir)) Directory.Delete(_dir, recursive: true);
    }

    [Fact]
    public void GeneratedFilesAreReadBackWithOnlyInjectedIssues()
    {
        var files = SampleDataGenerator.Generate(_dir, Month);
        var (result, fileCount) = BranchReportReader.ReadDirectory(_dir);

        Assert.Equal(3, files.Count);
        Assert.Equal(3, fileCount);
        // 每份檔案刻意放一筆金額 N/A，其他手打問題要能自動修正
        Assert.Equal(3, result.Issues.Count);
        Assert.All(result.Issues, i => Assert.Equal("金額無法辨識", i.Problem));

        Assert.Equal(SampleDataGenerator.Branches.Order(), result.Records.Select(r => r.Branch).Distinct().Order());
        Assert.All(result.Records, r =>
        {
            Assert.Equal(Month.Year, r.Date.Year);
            Assert.Equal(Month.Month, r.Date.Month);
            Assert.True(SampleDataGenerator.Menu.ContainsKey(r.Item), $"品項未清理：'{r.Item}'");
            Assert.Equal(r.Quantity * SampleDataGenerator.Menu[r.Item], r.Amount);
        });
    }

    [Fact]
    public void SameSeedProducesSameData()
    {
        var a = Path.Combine(_dir, "a");
        var b = Path.Combine(_dir, "b");
        SampleDataGenerator.Generate(a, Month, seed: 7);
        SampleDataGenerator.Generate(b, Month, seed: 7);

        Assert.Equal(BranchReportReader.ReadDirectory(a).Result.Records, BranchReportReader.ReadDirectory(b).Result.Records);
    }

    [Fact]
    public void WritesSummaryWorkbook()
    {
        SampleDataGenerator.Generate(_dir, Month);
        var (result, fileCount) = BranchReportReader.ReadDirectory(_dir);
        var output = Path.Combine(_dir, "out", "彙總報表.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);

        SummaryReportWriter.Write(output, result.Records, result.Issues, fileCount, new DateTime(2026, 11, 7));

        using var wb = new XLWorkbook(output);
        Assert.Equal(["總覽", "品項排行", "合併明細", "異常資料"], wb.Worksheets.Select(w => w.Name));

        // 總覽的合計列公式 = 所有有效資料金額
        var overview = wb.Worksheet("總覽");
        var totalRow = overview.Column(1).CellsUsed().First(c => c.GetString() == "合計").Address.RowNumber;
        Assert.Equal((double)result.Records.Sum(r => r.Amount), overview.Cell(totalRow, 4).Value.GetNumber());

        Assert.Equal(result.Records.Count + 1, wb.Worksheet("合併明細").LastRowUsed()!.RowNumber());
        Assert.Equal(SampleDataGenerator.Menu.Count + 1, wb.Worksheet("品項排行").LastRowUsed()!.RowNumber());
        Assert.Equal(result.Issues.Count + 1, wb.Worksheet("異常資料").LastRowUsed()!.RowNumber());
    }

    [Fact]
    public void OmitsIssuesSheetWhenDataIsClean()
    {
        var output = Path.Combine(_dir, "clean.xlsx");
        Directory.CreateDirectory(_dir);
        SummaryReportWriter.Write(output, [new("台北店", Month, "珍珠奶茶", 1, 55)], [], 1, DateTime.Now);

        using var wb = new XLWorkbook(output);
        Assert.DoesNotContain(wb.Worksheets, w => w.Name == "異常資料");
    }
}
