using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;

namespace PdfParser
{
    public static class GetPdfFields
    {
        public static string GetRectangle(PdfReader reader, int pageNumber, float llx, float lly, float urx, float ury)
        {
            iTextSharp.text.Rectangle rect = new iTextSharp.text.Rectangle(llx, lly, urx, ury);
            RenderFilter[] renderFilter = new RenderFilter[1];
            renderFilter[0] = new RegionTextRenderFilter(rect);
            ITextExtractionStrategy textExtractionStrategy = new FilteredTextRenderListener(new LocationTextExtractionStrategy(), renderFilter);
            return PdfTextExtractor.GetTextFromPage(reader, pageNumber, textExtractionStrategy);
        }

        public static string GetName(string content)
        {
            string start = "Tanımı";
            string end = "Site";
            int Start = 0, End = 0;
            if (content.Contains(start) && content.Contains(end))
            {
                Start = content.IndexOf(start, 0) + start.Length;
                End = content.IndexOf(end, Start);
                return content.Substring(Start, End - Start).Trim();
            }
            else
                return string.Empty;
        }

        public static string GetGate(string content)
        {
            string start = "Kapı";
            string end = "GLOBAL";
            int Start = 0, End = 0;
            if (content.Contains(start) && content.Contains(end))
            {
                Start = content.IndexOf(start, 0) + start.Length;
                End = content.IndexOf(end, Start);
                return content.Substring(Start, End - Start).Trim();
            }
            else
                return string.Empty;
        }

        public static string GetCardNo(string content)
        {
            string start = "İşveren no";
            string end = "POİNTER";
            int Start = 0, End = 0;
            if (content.Contains(start) && content.Contains(end))
            {
                Start = content.IndexOf(start, 0) + start.Length;
                End = content.IndexOf(end, Start);
                return content.Substring(Start, End - Start).Trim();
            }
            else
                return string.Empty;
        }

        public static DateTime GetDate(string content)
        {
            string start = "Host Geçmişi";
            string end = "Advisor MASTER";
            int Start = 0, End = 0;

            if (content.Contains(start) && content.Contains(end))
            {
                Start = content.IndexOf(start, 0) + start.Length;
                End = content.IndexOf(end, Start);
                
            }
            return DateTime.Parse(content.Substring(Start, End - Start).Trim());
        }
    }
}
