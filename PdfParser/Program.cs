using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;

namespace PdfParser
{
    public class Program
    {
        public static string excelSavePath = @"C:\Users\Mehmet Serkan Ekinci\Desktop\dwg2.xlsx";
        public static string pdfPath = @"C:\\Users\\Mehmet Serkan Ekinci\\Downloads\\ocak-personel2.pdf";
        static void Main(string[] args)
        {
            List<Informations> employeeList = PdfParser.PdfText(pdfPath);
            List<XmlModel> xmlModels = PdfParser.ParseToXmlModel(employeeList);
            PdfParser.ListToExcel(xmlModels, excelSavePath);
        }
    }
}
