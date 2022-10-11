using FinancialReportTool.DataModels;

namespace FinancialReportTool.Interfaces;
public interface IReportSaver
{
    public void Save(ReportModel report);
}