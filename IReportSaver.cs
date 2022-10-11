using FinancialReportTool.DataModels;

namespace FinancialReportTool;
public interface IReportSaver
{
    public void Save(ReportModel report);
}