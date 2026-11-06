using Bogus;
using ClosedXML.Excel;

namespace ExcelAutomationDemo;

/// <summary>
/// 用 Bogus 產生假的「分店月報」Excel，模擬各分店人員手動維護的檔案：
/// 有標題列、空白列、手打的合計列，以及少量格式不一致的資料。
/// </summary>
public static class SampleDataGenerator
{
    public static readonly string[] Branches = ["台北店", "台中店", "高雄店"];

    /// <summary>飲料店品項與單價（假資料）。</summary>
    public static readonly IReadOnlyDictionary<string, decimal> Menu = new Dictionary<string, decimal>
    {
        ["珍珠奶茶"] = 55, ["紅茶拿鐵"] = 60, ["四季春青茶"] = 35, ["冬瓜檸檬"] = 50,
        ["芒果冰沙"] = 75, ["鮮奶綠茶"] = 55, ["黑糖鮮奶"] = 70, ["檸檬綠茶"] = 45,
    };

    public static string FileNameFor(string branch, DateOnly month) => $"分店月報_{branch}_{month:yyyy-MM}.xlsx";

    /// <summary>在 <paramref name="directory"/> 產生每間分店一份月報，回傳檔案路徑。</summary>
    public static IReadOnlyList<string> Generate(string directory, DateOnly month, int seed = 42)
    {
        Directory.CreateDirectory(directory);
        var paths = new List<string>();
        for (var i = 0; i < Branches.Length; i++)
        {
            var path = Path.Combine(directory, FileNameFor(Branches[i], month));
            WriteBranchFile(path, Branches[i], month, new Faker { Random = new Randomizer(seed + i) });
            paths.Add(path);
        }
        return paths;
    }

    private static void WriteBranchFile(string path, string branch, DateOnly month, Faker faker)
    {
        var first = new DateOnly(month.Year, month.Month, 1);
        var days = DateTime.DaysInMonth(month.Year, month.Month);
        var items = Menu.Keys.ToArray();
        // 每間分店的熱門品項不同，排行才會有差異
        var raw = items.Select(_ => faker.Random.Float(0.3f, 1f)).ToArray();
        var weights = raw.Select(w => w / raw.Sum()).ToArray();

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("銷售明細");
        ws.Cell(1, 1).Value = $"{branch} {month:yyyy 年 M 月}銷售月報";
        ws.Range(1, 1, 1, 5).Merge().Style.Font.Bold = true;
        string[] headers = ["日期", "品項", "數量", "單價", "金額"];
        for (var c = 0; c < headers.Length; c++) ws.Cell(3, c + 1).Value = headers[c];
        ws.Range(3, 1, 3, headers.Length).Style.Font.Bold = true;

        var row = 4;
        for (var d = 0; d < days; d++)
        {
            var date = first.AddDays(d);
            var count = faker.Random.Int(3, 7);
            var picked = Enumerable.Range(0, count)
                .Select(_ => faker.Random.WeightedRandom(items, weights))
                .Distinct();
            foreach (var item in picked)
            {
                var qty = faker.Random.Int(5, 40);
                ws.Cell(row, 1).Value = date.ToDateTime(TimeOnly.MinValue);
                ws.Cell(row, 2).Value = item;
                ws.Cell(row, 3).Value = qty;
                ws.Cell(row, 4).Value = Menu[item];
                ws.Cell(row, 5).Value = qty * Menu[item];
                row++;
            }
            // 月中留一行空白，模擬人工編輯
            if (d == days / 2) row++;
        }
        var lastDataRow = row - 1;

        InjectTypicalMistakes(ws, faker, lastDataRow);

        // 人工加的合計列（彙總時要略過，不然會重複計算）
        ws.Cell(row, 2).Value = "合計";
        ws.Cell(row, 3).FormulaA1 = $"SUM(C4:C{lastDataRow})";
        ws.Cell(row, 5).FormulaA1 = $"SUM(E4:E{lastDataRow})";

        ws.Column(1).Style.DateFormat.Format = "yyyy/mm/dd";
        ExcelLayout.FitColumns(ws, 5, minWidth: 8);
        wb.SaveAs(path);
    }

    /// <summary>各放一個常見的手打錯誤：品項多空白、數字存成文字、金額填 N/A。</summary>
    private static void InjectTypicalMistakes(IXLWorksheet ws, Faker faker, int lastDataRow)
    {
        var dataRows = Enumerable.Range(4, lastDataRow - 3).Where(r => !ws.Cell(r, 2).IsEmpty()).ToList();
        var rows = faker.Random.ListItems(dataRows, 3);

        var item = ws.Cell(rows[0], 2);
        item.Value = $" {item.GetString()}  ";                 // 可自動修正

        var qty = ws.Cell(rows[1], 3);
        qty.Value = qty.GetValue<int>().ToString();             // 可自動修正
        qty.Style.NumberFormat.Format = "@";

        ws.Cell(rows[2], 5).Value = "N/A";                     // 無法判斷，列入異常
    }
}
