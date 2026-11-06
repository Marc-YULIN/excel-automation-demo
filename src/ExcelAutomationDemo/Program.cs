using System.Diagnostics;
using System.Globalization;
using System.Text;
using ExcelAutomationDemo;

Console.OutputEncoding = Encoding.UTF8;

const string Usage = """
    用法：
      ExcelAutomationDemo                     一鍵示範：產生範例月報 → 彙總
      ExcelAutomationDemo generate [選項]     只產生範例月報
      ExcelAutomationDemo merge [選項]        只彙總資料夾內的月報

    選項：
      --dir <資料夾>     月報所在資料夾，預設 sample-input
      --output <檔名>    彙總報表檔名，預設 彙總報表.xlsx
      --month <yyyy-MM>  範例資料月份，預設上個月
      --seed <數字>      亂數種子，相同種子產生相同資料，預設 42
      -h, --help         顯示說明
    """;

var command = args.Length > 0 && !args[0].StartsWith('-') ? args[0] : "demo";
var options = args.Length > 0 && !args[0].StartsWith('-') ? args[1..] : args;

string dir = "sample-input", output = "彙總報表.xlsx";
var lastMonth = DateOnly.FromDateTime(DateTime.Today).AddMonths(-1);
var month = new DateOnly(lastMonth.Year, lastMonth.Month, 1);
var seed = 42;

for (var i = 0; i < options.Length; i++)
{
    var hasValue = i + 1 < options.Length;
    switch (options[i])
    {
        case "-h" or "--help":
            Console.WriteLine(Usage);
            return 0;
        case "--dir" when hasValue:
            dir = options[++i];
            break;
        case "--output" when hasValue:
            output = options[++i];
            break;
        case "--month" when hasValue && DateOnly.TryParseExact(options[i + 1], "yyyy-MM",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out var m):
            month = m;
            i++;
            break;
        case "--seed" when hasValue && int.TryParse(options[i + 1], out var s):
            seed = s;
            i++;
            break;
        default:
            return Fail($"無法辨識的參數：{options[i]}");
    }
}

return command switch
{
    "demo" => Generate() == 0 ? Merge() : 1,
    "generate" => Generate(),
    "merge" => Merge(),
    _ => Fail($"無法辨識的指令：{command}"),
};

int Generate()
{
    var files = SampleDataGenerator.Generate(dir, month, seed);
    Console.WriteLine($"已產生 {files.Count} 份 {month:yyyy-MM} 範例月報：");
    foreach (var f in files) Console.WriteLine($"  {Path.GetFullPath(f)}");
    Console.WriteLine();
    return 0;
}

int Merge()
{
    if (!Directory.Exists(dir))
        return Fail($"找不到資料夾：{Path.GetFullPath(dir)}");

    var sw = Stopwatch.StartNew();
    var (result, fileCount) = BranchReportReader.ReadDirectory(dir);
    Console.WriteLine($"讀取 {fileCount} 個檔案，有效資料 {result.Records.Count:N0} 筆，異常 {result.Issues.Count} 筆");

    if (result.Records.Count == 0)
    {
        foreach (var issue in result.Issues) Console.WriteLine($"  ! {issue.File}：{issue.Problem}");
        Console.Error.WriteLine("沒有可彙總的資料，不產生報表。");
        return 2;
    }

    var outputPath = Path.GetFullPath(output);
    try
    {
        SummaryReportWriter.Write(outputPath, result.Records, result.Issues, fileCount, DateTime.Now);
    }
    catch (IOException)
    {
        return Fail($"無法寫入 {outputPath}，檔案可能正在 Excel 中開啟，請先關閉。");
    }

    foreach (var b in ReportCalculator.SummarizeBranches(result.Records))
        Console.WriteLine($"  {b.Branch}：NT${b.Amount:N0}（{b.Share:P1}）");
    foreach (var issue in result.Issues)
        Console.WriteLine($"  ! {issue.File} 第 {issue.Row} 列：{issue.Problem}");

    Console.WriteLine();
    Console.WriteLine($"完成，耗時 {sw.Elapsed.TotalSeconds:0.0} 秒");
    Console.WriteLine($"彙總報表：{outputPath}");
    return 0;
}

static int Fail(string message)
{
    Console.Error.WriteLine(message);
    Console.WriteLine(Usage);
    return 1;
}
