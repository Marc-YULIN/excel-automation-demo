using ClosedXML.Excel;

namespace ExcelAutomationDemo;

internal static class ExcelLayout
{
    /// <summary>
    /// 依內容設定欄寬。ClosedXML 的 AdjustToContents 會把中文字估成半形寬度，
    /// 所以自己算：全形字算 2、半形字算 1，略過合併儲存格（標題列）。
    /// </summary>
    public static void FitColumns(IXLWorksheet ws, int columnCount, double minWidth, double maxWidth = 60)
    {
        for (var c = 1; c <= columnCount; c++)
        {
            var widest = ws.Column(c).CellsUsed()
                .Where(cell => !cell.IsMerged())
                .Select(cell => DisplayWidth(cell.GetFormattedString()))
                .DefaultIfEmpty(0)
                .Max();
            ws.Column(c).Width = Math.Clamp(widest + 3, minWidth, maxWidth);
        }
    }

    public static int DisplayWidth(string text) => text.Sum(ch => ch >= '⺀' ? 2 : 1);
}
