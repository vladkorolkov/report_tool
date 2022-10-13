using FinancialReportTool.DataModels;

namespace FinancialReportTool.Interfaces;
public interface IReportHandler
{
    public List<ReportModel> HandleReport(QueryModel query);
}