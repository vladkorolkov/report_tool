using FinancialReportTool.DataModels;
namespace FinancialReportTool;
public interface IQueryService
{
    public QueryModel GetQuery(string[] args);
}