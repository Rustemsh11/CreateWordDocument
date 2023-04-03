
using Microsoft.Office.Interop.Word;

namespace WordApp
{
    public class WordDocument
    {
        public string SuplyerName { get; set; }
        public string SuplyerINN { get; set; }
        public string SuplyerKPP { get; set; }
        public string CompanyName { get; set; }
        public string CompanyINN { get; set; }
        public string CompanyKPP { get; set; }
        public string SnabName { get; set; }
        public string SnabPhone { get; set; }
        public string Email { get; set; }


        public void Create()
        {
            Application application = new Application();
            Document document = application.Documents.Add(Visible: true);
            Paragraph paragraph = document.Paragraphs.Add();
            document.Paragraphs.Format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            document.Paragraphs.Format.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;

            Range range = document.Paragraphs[1].Range;
            range.Text = string.Format("Кому: {0}\vИнн: {1} КПП: {2}\v\v От: {3}\vИнн: {4} КПП: {5}\vКонтактное лицо: {6}\vТелефон: {7}\v Email: {8}"
                , SuplyerName, SuplyerINN, SuplyerKPP, CompanyName, CompanyINN, CompanyKPP, SnabName, SnabPhone,Email);
            range.Font.Name = "Times New Roman";
            range.Font.Size = 14;

            Paragraph paragraph1 = document.Paragraphs.Add();
            document.Paragraphs[2].Format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            Range range2 = document.Paragraphs[2].Range;
            range2.Text = "Запрос о предоставлении ценовой информации";
            range2.Font.Name = "Times New Roman";
            range2.Font.Size = 20;
            range2.Font.Bold = 0;            
            document.SaveAs(@"C:\Users\агроном\Desktop\WordTest\test.docx");
            document.Close();

            application.Quit();
        }
    }
}