using FinancialReportTool.DataModels;

namespace FinancialReportTool.Interfaces;
public interface IReportHandler
{
    public List<ReportModel> Read (QueryModel query);
}