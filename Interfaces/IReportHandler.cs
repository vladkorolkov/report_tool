using FinancialReportTool.DataModels;

namespace FinancialReportTool;
public interface IReportHandler
{
    public ReportModel Edit(QueryModel query);
}