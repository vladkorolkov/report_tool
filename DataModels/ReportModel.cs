namespace FinancialReportTool.DataModels;
public class ReportModel
{
    public int N { get; set; }
    public string Period { get; set; }
    public string Platform { get; set; }
    public string TypeOfRights { get; set; }
    public string Territory { get; set; }
    public string ContentType { get; set; }
    public string UsageType { get; set; }
    public string ArtistName { get; set; }
    public string Track { get; set; }
    public string Album { get; set; }
    public string LyricsAuthor { get; set; }
    public string Composer { get; set; }
    public double CopyrightShare { get; set; }
    public double RelatedCopyrightShare { get; set; }
    public string Isrc { get; set; }
    public string Upc { get; set; }
    public string Copyright { get; set; }
    public int Listens { get; set; }
    public decimal NdaCopyrightsRecieved { get; set; }
    public decimal NdaRelatedCopyrightsRecieved { get; set; }
    public double LabelShare { get; set; }
    public decimal LabelCopyrightsRecieved { get; set; }
    public decimal LabelRelatedCopyrightsRecieved { get; set; }
    public decimal Total { get; set; }

}
