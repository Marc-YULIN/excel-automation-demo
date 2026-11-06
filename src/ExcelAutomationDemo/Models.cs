namespace ExcelAutomationDemo;

/// <summary>一筆銷售紀錄（從分店月報讀進來、清理過的資料）。</summary>
public sealed record SaleRecord(string Branch, DateOnly Date, string Item, int Quantity, decimal Amount);

/// <summary>讀檔時發現、無法自動修正的資料問題。</summary>
/// <param name="File">來源檔名</param>
/// <param name="Row">工作表列號；整個檔案的問題為 null</param>
/// <param name="Problem">問題說明</param>
/// <param name="RawContent">原始儲存格內容，方便回頭找人確認</param>
public sealed record DataIssue(string File, int? Row, string Problem, string RawContent);

/// <summary>分店小計。</summary>
public sealed record BranchSummary(
    string Branch, int Transactions, int Quantity, decimal Amount, decimal Share, string TopItem);

/// <summary>品項排行的一列。</summary>
public sealed record ItemRanking(
    int Rank, string Item, int Quantity, decimal Amount, decimal Share,
    IReadOnlyDictionary<string, decimal> AmountByBranch);
