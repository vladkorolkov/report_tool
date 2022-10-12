using FinancialReportTool.DataModels;
using FinancialReportTool.Interfaces;
using OfficeOpenXml;

namespace FinancialReportTool.Services;
public class ReportHandler : IReportHandler
{
    private const int UnsedRowsNumber = 8;
    public List<ReportModel> Edit(QueryModel query)
    {
        
        //for free usage
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var rootSheet = new ExcelPackage(query.Path).Workbook.Worksheets[1];
        rootSheet.DeleteRow(1, 5, true);
        var allRowsCount = rootSheet.Dimension.Rows;
        var rowsWithValues = allRowsCount - UnsedRowsNumber ;
        var result = new List<ReportModel>();
        for (int currentRowNumber = 2; currentRowNumber < rowsWithValues; currentRowNumber++)
        {
            var currentRow = new ReportModel();
            //currentRow.N = int.Parse(rootSheet.Cells[$"A{currentRowNumber}"].Value?.ToString());
            currentRow.Period = rootSheet.Cells[$"B{currentRowNumber}"].Value.ToString()??"";
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


        var exPkg = new ExcelPackage($"{query.Artist}.xlsx");
        
        var sheet = exPkg.Workbook.Worksheets.Add("Отчет по прослушиваниям");

        sheet.Cells["A1"].Value = "Трек";
        sheet.Cells["B1"].Value = "Альбом";
        sheet.Cells["C1"].Value = "Плошадка";
        sheet.Cells["D1"].Value = "Прослушивания";
        sheet.Cells["E1"].Value = "Территория";
        sheet.Cells["F1"].Value = "Период";
        sheet.Cells["G1"].Value = "Вознагрждение, руб.";
        sheet.Cells["H1"].Value = "Итого, руб.";

        var totalRows = result.Count();
        decimal totalCashEarnded = 0;
        for (int currentRow=2; currentRow <= totalRows; currentRow++)
        {
           
                sheet.Cells[$"A{currentRow}"].Value = result[currentRow-2].Track;
                sheet.Cells[$"B{currentRow}"].Value = result[currentRow-2].Album;
                sheet.Cells[$"C{currentRow}"].Value = result[currentRow-2].Platform;
                sheet.Cells[$"D{currentRow}"].Value = result[currentRow-2].Listens;
                sheet.Cells[$"E{currentRow}"].Value = result[currentRow-2].Territory;
                sheet.Cells[$"F{currentRow}"].Value = result[currentRow-2].Period;
                sheet.Cells[$"G{currentRow}"].Value = result[currentRow-2].Total;
                totalCashEarnded += result[currentRow-2].Total;
            
        }
        sheet.Cells[$"H1"].Value = totalCashEarnded;

        exPkg.Save();
        return result;
    }

    private double ParseDouble(string value)
    {
        double result;
        if(double.TryParse(value, out result))       
            return result;      
        return 0;        
    }
    private decimal ParseDecimal(string value)
    {
        decimal result;
        if(decimal.TryParse(value, out result))
            return result;
        return 0;
    }
}