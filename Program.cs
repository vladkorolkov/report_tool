using FinancialReportTool.Services;
namespace FinancialReportTool;
class Program
{
    private static void Main(string[] args)
    {
        try
        {
            var query = new QueryService(args).GetQuery();
            var reportService = new ReportHandler();
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
            Console.WriteLine("Не указан один из обязательных параметров:\n1. Имя исходного файла отчета\n2. Имя артиста или название группы. Если название группы состоит из двух и более слов, укажите название в кавычках.\n\nПример:\nreportfile.xlsx Metallica\nreportfile.xlsx \"System of a Dowm\"\n\nЧтобы продолжить, закройте это окно и запустите программу снова.");
            Console.ReadLine();
            return;
        }
    }
}
