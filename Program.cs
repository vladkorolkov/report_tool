namespace FinancialReportTool;
class Program
{
    private static void Main(string[] args)
    {
        
#if DEBUG
        args = new[] { "test.xlsx -Alpin" };
#endif
        var query = new QueryService(args).GetQuery();
        
        Console.WriteLine($"First parameter: {query.Path}");
        Console.WriteLine($"Second parameter: {query.Artist}");
        var report = new Services.ReportHandler();
        var result = report.Edit(query);
        Console.ReadLine();
    }
} 
