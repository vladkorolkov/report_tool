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
        var result = report.Read(query);
        var saved = report.Save(result, query.Artist);
        if (saved)
            Console.WriteLine("Отчет успешно создан в этой же папке.\nЧтобы выйти, нажмите ctrl-c");
        else
            Console.WriteLine("Что-то пошло не так.\nВ папке, из которой было запущено приложение создан файл логов, отправьте его в телегу @vlad_korolkov");
        Console.ReadLine();
    }
}
