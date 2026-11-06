using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace ExcelAutomationDemo;

/// <summary>
/// 讀取分店月報。不依賴固定格式：自動找標題列、欄位順序可不同，
/// 並清理常見的人工輸入問題；無法判斷的資料列入 <see cref="DataIssue"/>，不會默默算錯。
/// </summary>
public static partial class BranchReportReader
{
    public sealed record ReadResult(IReadOnlyList<SaleRecord> Records, IReadOnlyList<DataIssue> Issues);

    private static readonly string[] RequiredHeaders = ["日期", "品項", "數量", "金額"];
    private static readonly string[] TotalMarkers = ["合計", "小計", "總計"];
    private static readonly string[] DateFormats = ["yyyy/M/d", "yyyy-M-d", "yyyy.M.d", "yyyyMMdd"];
    private const int HeaderSearchRows = 20;

    /// <summary>讀取資料夾內所有 .xlsx（略過 Excel 暫存檔 ~$*.xlsx），依檔名排序。</summary>
    public static (ReadResult Result, int FileCount) ReadDirectory(string directory)
    {
        var files = Directory.GetFiles(directory, "*.xlsx")
            .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.Ordinal))
            .Order(StringComparer.Ordinal)
            .ToList();

        var records = new List<SaleRecord>();
        var issues = new List<DataIssue>();
        foreach (var file in files)
        {
            var result = Read(file);
            records.AddRange(result.Records);
            issues.AddRange(result.Issues);
        }
        return (new ReadResult(records, issues), files.Count);
    }

    public static ReadResult Read(string path)
    {
        var fileName = Path.GetFileName(path);
        try
        {
            using var wb = new XLWorkbook(path);
            return Read(wb, fileName);
        }
        catch (IOException ex)
        {
            return new ReadResult([], [new DataIssue(fileName, null, "檔案無法開啟（可能正在 Excel 中開啟）", ex.Message)]);
        }
    }

    public static ReadResult Read(XLWorkbook wb, string fileName)
    {
        var branch = BranchNameFromFile(fileName);
        var ws = wb.Worksheets.FirstOrDefault(w => w.Name == "銷售明細") ?? wb.Worksheet(1);
        var records = new List<SaleRecord>();
        var issues = new List<DataIssue>();

        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
        if (FindHeader(ws, lastRow) is not var (headerRow, col))
        {
            issues.Add(new DataIssue(fileName, null,
                $"找不到標題列（需包含：{string.Join("、", RequiredHeaders)}）", ""));
            return new ReadResult(records, issues);
        }

        for (var r = headerRow + 1; r <= lastRow; r++)
        {
            var dateCell = ws.Cell(r, col["日期"]);
            var itemCell = ws.Cell(r, col["品項"]);
            var qtyCell = ws.Cell(r, col["數量"]);
            var amountCell = ws.Cell(r, col["金額"]);
            IXLCell[] cells = [dateCell, itemCell, qtyCell, amountCell];

            if (cells.All(c => string.IsNullOrWhiteSpace(c.GetFormattedString())))
                continue; // 空白列

            var item = WhitespaceRegex().Replace(itemCell.GetFormattedString(), " ").Trim();
            if (IsTotalRow(item) || IsTotalRow(dateCell.GetFormattedString()))
                continue; // 人工合計列，略過避免重複計算

            var problems = new List<string>();
            var date = ParseDate(dateCell);
            if (date is null) problems.Add("日期無法辨識");
            if (item.Length == 0) problems.Add("品項空白");
            var qty = ParseNumber(qtyCell);
            if (qty is null || qty != Math.Floor(qty.Value)) problems.Add("數量無法辨識");
            var amount = ParseNumber(amountCell);
            if (amount is null) problems.Add("金額無法辨識");

            if (problems.Count > 0)
            {
                var raw = string.Join(" | ", cells.Select(c => c.GetFormattedString()));
                issues.Add(new DataIssue(fileName, r, string.Join("、", problems), raw));
                continue;
            }

            records.Add(new SaleRecord(branch, date!.Value, item, (int)qty!.Value, amount!.Value));
        }

        return new ReadResult(records, issues);
    }

    /// <summary>「分店月報_台北店_2026-10.xlsx」→「台北店」；格式不符就用檔名。</summary>
    public static string BranchNameFromFile(string fileName)
    {
        var name = Path.GetFileNameWithoutExtension(fileName);
        var match = BranchFileRegex().Match(name);
        return match.Success ? match.Groups["branch"].Value : name;
    }

    private static (int Row, Dictionary<string, int> Columns)? FindHeader(IXLWorksheet ws, int lastRow)
    {
        for (var r = 1; r <= Math.Min(lastRow, HeaderSearchRows); r++)
        {
            var columns = new Dictionary<string, int>();
            foreach (var cell in ws.Row(r).CellsUsed())
            {
                var text = cell.GetFormattedString().Trim();
                if (RequiredHeaders.Contains(text)) columns.TryAdd(text, cell.Address.ColumnNumber);
            }
            if (columns.Count == RequiredHeaders.Length) return (r, columns);
        }
        return null;
    }

    private static bool IsTotalRow(string text) =>
        TotalMarkers.Any(m => text.Contains(m, StringComparison.Ordinal));

    private static DateOnly? ParseDate(IXLCell cell)
    {
        var value = cell.Value;
        if (value.IsDateTime) return DateOnly.FromDateTime(value.GetDateTime());
        // 沒有套日期格式的日期，存的是 Excel 序號
        if (value.IsNumber && value.GetNumber() is > 1 and < 2958466)
            return DateOnly.FromDateTime(DateTime.FromOADate(value.GetNumber()));
        if (value.IsText &&
            DateTime.TryParseExact(value.GetText().Trim(), DateFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var parsed))
            return DateOnly.FromDateTime(parsed);
        return null;
    }

    private static decimal? ParseNumber(IXLCell cell)
    {
        var value = cell.Value;
        if (value.IsNumber) return (decimal)value.GetNumber();
        if (!value.IsText) return null;
        // 存成文字的數字，例如 "12"、"1,200"、"NT$ 1,200"
        var text = value.GetText().Replace("NT$", "").Replace("$", "").Replace(",", "").Trim();
        return decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var n) ? n : null;
    }

    [GeneratedRegex(@"^分店月報_(?<branch>.+?)_\d{4}-\d{2}$")]
    private static partial Regex BranchFileRegex();

    [GeneratedRegex(@"\s+")]
    private static partial Regex WhitespaceRegex();
}
