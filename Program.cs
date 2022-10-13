namespace FinancialReportTool;
class Program
{
    private static void Main(string[] args)
    {
        try
        {
            var query = new QueryService(args).GetQuery();
            Console.WriteLine($"First parameter: {query.Path}");
            Console.WriteLine($"Second parameter: {query.Artist}");
            var reportService = new Services.ReportHandler();
            var result = reportService.HandleReport(query);
            var saved = reportService.Save(result, query.Artist);
            if (saved)
                Console.WriteLine("\n\nОтчет успешно создан в этой же папке.\nЧтобы выйти, нажмите ctrl-c");
            else
                Console.WriteLine("Что-то пошло не так.\nВ папке, из которой было запущено приложение создан файл логов, отправьте его в телегу @vlad_korolkov");
            Console.ReadLine();
        }
        catch (IndexOutOfRangeException ex)
        {
            Console.WriteLine("Не указан один из обязательных параметров:\n1. Имя исходного файла отчета\n2. Имя артиста(название группы)\nПример:\nreportfile.xlsx Metallica\n\nЧтобы продолжить, закройте это окно и запустите программу снова.");
            Console.ReadLine();
            return;
        }



    }
}
