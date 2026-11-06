namespace ExcelAutomationDemo;

/// <summary>彙總計算（純計算，不碰檔案，方便測試）。</summary>
public static class ReportCalculator
{
    /// <summary>各分店小計，依總金額由高到低。</summary>
    public static IReadOnlyList<BranchSummary> SummarizeBranches(IReadOnlyList<SaleRecord> records)
    {
        var total = records.Sum(r => r.Amount);
        return records
            .GroupBy(r => r.Branch)
            .Select(g => new BranchSummary(
                Branch: g.Key,
                Transactions: g.Count(),
                Quantity: g.Sum(r => r.Quantity),
                Amount: g.Sum(r => r.Amount),
                Share: total == 0 ? 0 : g.Sum(r => r.Amount) / total,
                TopItem: g.GroupBy(r => r.Item)
                    .OrderByDescending(i => i.Sum(r => r.Amount))
                    .ThenBy(i => i.Key, StringComparer.Ordinal)
                    .First().Key))
            .OrderByDescending(b => b.Amount)
            .ThenBy(b => b.Branch, StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>品項排行：依總金額排序，同金額再比數量。</summary>
    public static IReadOnlyList<ItemRanking> RankItems(IReadOnlyList<SaleRecord> records)
    {
        var total = records.Sum(r => r.Amount);
        var branches = records.Select(r => r.Branch).Distinct().ToList();

        return records
            .GroupBy(r => r.Item)
            .Select(g => (
                Item: g.Key,
                Quantity: g.Sum(r => r.Quantity),
                Amount: g.Sum(r => r.Amount),
                ByBranch: branches.ToDictionary(
                    b => b, b => g.Where(r => r.Branch == b).Sum(r => r.Amount))))
            .OrderByDescending(x => x.Amount)
            .ThenByDescending(x => x.Quantity)
            .ThenBy(x => x.Item, StringComparer.Ordinal)
            .Select((x, i) => new ItemRanking(
                Rank: i + 1,
                Item: x.Item,
                Quantity: x.Quantity,
                Amount: x.Amount,
                Share: total == 0 ? 0 : x.Amount / total,
                AmountByBranch: x.ByBranch))
            .ToList();
    }
}
