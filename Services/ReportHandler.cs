using FinancialReportTool.DataModels;
using FinancialReportTool.Interfaces;
using OfficeOpenXml;

namespace FinancialReportTool.Services;
public class ReportHandler : IReportHandler
{
    private const int UnsedRowsNumberFromTop = 5;
    private const int UnsedRowsNumberFromBottom = 8;
    private List<ReportModel> Read(QueryModel query)
    {
        //for free usage epplus lib
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var rootSheet = new ExcelPackage(query.Path).Workbook.Worksheets[1];
        rootSheet.DeleteRow(1, UnsedRowsNumberFromTop);

        var allRowsCount = rootSheet.Dimension.Rows;
        var rowsWithValues = allRowsCount - UnsedRowsNumberFromBottom;

        var result = new List<ReportModel>();
        for (int currentRowNumber = 2; currentRowNumber < rowsWithValues; currentRowNumber++)
        {
            var currentRow = new ReportModel();
            //currentRow.N = int.Parse(rootSheet.Cells[$"A{currentRowNumber}"].Value?.ToString());
            currentRow.Period = rootSheet.Cells[$"B{currentRowNumber}"].Value.ToString() ?? "";
            currentRow.Platform = rootSheet.Cells[$"C{currentRowNumber}"].Value.ToString();
            currentRow.TypeOfRights = rootSheet.Cells[$"D{currentRowNumber}"].Value.ToString();
            currentRow.Territory = rootSheet.Cells[$"E{currentRowNumber}"].Value.ToString();
            currentRow.ContentType = rootSheet.Cells[$"F{currentRowNumber}"].Value.ToString();
            currentRow.UsageType = rootSheet.Cells[$"G{currentRowNumber}"].Value.ToString();
            currentRow.ArtistName = rootSheet.Cells[$"H{currentRowNumber}"].Value.ToString();
            currentRow.Track = rootSheet.Cells[$"I{currentRowNumber}"].Value.ToString();
            currentRow.Album = rootSheet.Cells[$"J{currentRowNumber}"].Value.ToString();
            currentRow.LyricsAuthor = rootSheet.Cells[$"K{currentRowNumber}"].Value?.ToString();
            currentRow.Composer = rootSheet.Cells[$"L{currentRowNumber}"].Value?.ToString();
            currentRow.CopyrightShare = ParseDouble(rootSheet.Cells[$"M{currentRowNumber}"].Value.ToString());
            currentRow.RelatedCopyrightShare = ParseDouble(rootSheet.Cells[$"N{currentRowNumber}"].Value.ToString());
            currentRow.Isrc = rootSheet.Cells[$"O{currentRowNumber}"].Value.ToString();
            currentRow.Upc = rootSheet.Cells[$"P{currentRowNumber}"].Value?.ToString();
            currentRow.Copyright = rootSheet.Cells[$"Q{currentRowNumber}"].Value?.ToString();
            currentRow.Listens = int.Parse(rootSheet.Cells[$"R{currentRowNumber}"].Value.ToString());
            currentRow.NdaCopyrightsRecieved = ParseDecimal(rootSheet.Cells[$"S{currentRowNumber}"].Value.ToString());
            currentRow.NdaRelatedCopyrightsRecieved = ParseDecimal(rootSheet.Cells[$"T{currentRowNumber}"].Value?.ToString()); ;
            currentRow.LabelShare = ParseDouble(rootSheet.Cells[$"U{currentRowNumber}"].Value.ToString());
            currentRow.LabelCopyrightsRecieved = ParseDecimal(rootSheet.Cells[$"V{currentRowNumber}"].Value.ToString());
            currentRow.LabelRelatedCopyrightsRecieved = ParseDecimal(rootSheet.Cells[$"W{currentRowNumber}"].Value.ToString());
            currentRow.Total = ParseDecimal(rootSheet.Cells[$"X{currentRowNumber}"].Value.ToString());
            currentRow.Total = Math.Round(currentRow.Total, 2);
            result.Add(currentRow);
        }

        return result;
    }

    private double ParseDouble(string value)
    {
        double result;
        if (double.TryParse(value, out result))
            return result;
        return 0;
    }
    private decimal ParseDecimal(string value)
    {
        decimal result;
        if (decimal.TryParse(value, out result))
            return result;
        return 0;
    }

    public bool Save(List<ReportModel> report, string artistName)
    {
        
        
        var path = $"REPORTS/{artistName}.xlsx";
        Directory.CreateDirectory(Path.GetDirectoryName(path));
        
        if (File.Exists(path))
        {
            File.Delete(path);
        }
        var exPkg = new ExcelPackage(path);
        string sheetName = "?????????? ???? ????????????????????????????";
        var sheet = exPkg.Workbook.Worksheets.Add(sheetName);

        sheet.Cells["A1"].Value = "????????";
        sheet.Cells["B1"].Value = "????????????";
        sheet.Cells["C1"].Value = "????????????????";
        sheet.Cells["D1"].Value = "??????????????????????????";
        sheet.Cells["E1"].Value = "????????????????????";
        sheet.Cells["F1"].Value = "????????????";
        sheet.Cells["G1"].Value = "????????????????????????????, ??????.";
        sheet.Cells["H1"].Value = "??????????, ??????.";

        var totalRows = report.Count();
        decimal totalCashEarnded = 0;
        for (int currentRow = 2; currentRow <= totalRows; currentRow++)
        {
            sheet.Cells[$"A{currentRow}"].Value = report[currentRow - 2].Track;
            sheet.Cells[$"B{currentRow}"].Value = report[currentRow - 2].Album;
            sheet.Cells[$"C{currentRow}"].Value = report[currentRow - 2].Platform;
            sheet.Cells[$"D{currentRow}"].Value = report[currentRow - 2].Listens;
            sheet.Cells[$"E{currentRow}"].Value = report[currentRow - 2].Territory;
            sheet.Cells[$"F{currentRow}"].Value = report[currentRow - 2].Period;
            sheet.Cells[$"G{currentRow}"].Value = report[currentRow - 2].Total;
            totalCashEarnded += report[currentRow - 2].Total;
        }
        sheet.Cells[$"H1"].Value = totalCashEarnded;
        sheet.Cells["A1:H1"].Style.Font.Bold =true;
        sheet.Cells[$"A2:G{totalRows}"].AutoFitColumns();
        try
        {
            exPkg.Save();
        }
        catch (Exception ex)
        {
            string logsPath = "logs.txt";
            using (StreamWriter sw = File.CreateText(logsPath))
            {
                sw.WriteLine(DateTime.Now);
                sw.WriteLine(ex.Message);
                sw.WriteLine("======================================");
                sw.WriteLine(ex.StackTrace);
            }
            return false;
        }
        return true;
    }

    public List<ReportModel> HandleReport(QueryModel queryModel)
    {
        var reportInMemory = Read(queryModel);

        var query = reportInMemory.AsQueryable();
        if (!string.IsNullOrEmpty(queryModel.Artist))
            query = query.Where(u => u.ArtistName == queryModel.Artist);

        var filteredReport = query.ToList();
        return filteredReport;
    }
}