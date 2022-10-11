using FinancialReportTool.DataModels;

namespace FinancialReportTool.Interfaces;
public interface IReportHandler
{
    public ReportModel Edit(QueryModel query);
}