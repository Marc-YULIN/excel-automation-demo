namespace ExcelAutomationDemo.Tests;

public class ReportCalculatorTests
{
    private static readonly DateOnly Day = new(2026, 10, 1);

    private static readonly SaleRecord[] Records =
    [
        new("台北店", Day, "珍珠奶茶", 10, 550),
        new("台北店", Day, "紅茶拿鐵", 2, 120),
        new("台中店", Day, "紅茶拿鐵", 20, 1200),
        new("台中店", Day.AddDays(1), "珍珠奶茶", 1, 55),
        new("高雄店", Day, "四季春青茶", 3, 105),
    ];

    [Fact]
    public void SummarizesBranchesByAmount()
    {
        var branches = ReportCalculator.SummarizeBranches(Records);

        Assert.Equal(["台中店", "台北店", "高雄店"], branches.Select(b => b.Branch));
        var taichung = branches[0];
        Assert.Equal(2, taichung.Transactions);
        Assert.Equal(21, taichung.Quantity);
        Assert.Equal(1255m, taichung.Amount);
        Assert.Equal("紅茶拿鐵", taichung.TopItem);
        Assert.Equal(1m, branches.Sum(b => b.Share));
    }

    [Fact]
    public void RanksItemsWithPerBranchAmounts()
    {
        var ranking = ReportCalculator.RankItems(Records);

        Assert.Equal(["紅茶拿鐵", "珍珠奶茶", "四季春青茶"], ranking.Select(r => r.Item));
        Assert.Equal([1, 2, 3], ranking.Select(r => r.Rank));

        var tea = ranking[0];
        Assert.Equal(22, tea.Quantity);
        Assert.Equal(1320m, tea.Amount);
        Assert.Equal(120m, tea.AmountByBranch["台北店"]);
        Assert.Equal(1200m, tea.AmountByBranch["台中店"]);
        Assert.Equal(0m, tea.AmountByBranch["高雄店"]);
    }

    [Fact]
    public void BreaksAmountTiesByQuantity()
    {
        SaleRecord[] records =
        [
            new("A", Day, "甲", 1, 100),
            new("A", Day, "乙", 5, 100),
        ];
        Assert.Equal(["乙", "甲"], ReportCalculator.RankItems(records).Select(r => r.Item));
    }

    [Fact]
    public void HandlesEmptyInput()
    {
        Assert.Empty(ReportCalculator.SummarizeBranches([]));
        Assert.Empty(ReportCalculator.RankItems([]));
    }
}
