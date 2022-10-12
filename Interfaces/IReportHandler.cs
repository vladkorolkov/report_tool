using FinancialReportTool.DataModels;

namespace FinancialReportTool.Interfaces;
public interface IReportHandler
{
    public List<ReportModel> Edit(QueryModel query);
}