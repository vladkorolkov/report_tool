using FinancialReportTool.Interfaces;
using FinancialReportTool.DataModels;
public class QueryService : IQueryService
{
    private string[] _args;
    public QueryService(string[] args )
    {  
        _args = args;
    }
    public QueryModel GetQuery()
    {
        if (_args.Length < 0)
        {
            return new QueryModel();
        }
        var query = new QueryModel();

        if(_args[0] is not null & _args[1] is not null)
        {
            query.Path = _args[0];
            query.Artist = _args[1];
        }
      
        return query;
    }
}