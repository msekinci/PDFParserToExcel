using iTextSharp.text;
using iTextSharp.text.pdf;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PdfParser
{
        public static class PdfParser
        {
            public static float height = PageSize.A4.Height;
            public static float weight = PageSize.A4.Width;
            public static List<Informations> listOfInformations = new List<Informations>();
            public static List<string> partTexts = new List<string>();

            public static List<Informations> PdfText(string path)
            {
                PdfReader reader = new PdfReader(path);
                string text = string.Empty;
                for (int page = 1; page <= reader.NumberOfPages; page++)
                {
                    if (page == 1)
                    {
                        partTexts.Add(GetPdfFields.GetRectangle(reader, page, 0, height - 200, weight, height - 80));  // for part 1
                        partTexts.Add(GetPdfFields.GetRectangle(reader, page, 0, height - 380, weight, height - 200)); // for part 2
                        partTexts.Add(GetPdfFields.GetRectangle(reader, page, 0, height - 500, weight, height - 380)); // for part 3
                        partTexts.Add(GetPdfFields.GetRectangle(reader, page, 0, height - 660, weight, height - 500)); // for part 4
                        partTexts.Add(GetPdfFields.GetRectangle(reader, page, 0, height - 820, weight, height - 660)); // for part 5
                    }
                    else
                    {
                        partTexts.Add(GetPdfFields.GetRectangle(reader, page, 0, height - 190, weight, height - 50));  // for part 1
                        partTexts.Add(GetPdfFields.GetRectangle(reader, page, 0, height - 340, weight, height - 190)); // for part 2
                        partTexts.Add(GetPdfFields.GetRectangle(reader, page, 0, height - 480, weight, height - 340)); // for part 3
                        partTexts.Add(GetPdfFields.GetRectangle(reader, page, 0, height - 620, weight, height - 480)); // for part 4
                        partTexts.Add(GetPdfFields.GetRectangle(reader, page, 0, height - 750, weight, height - 620)); // for part 5
                    }
                }
                reader.Close();
                partTextAddToList(partTexts);
                return listOfInformations;
            }
            public static void partTextAddToList(List<string> partTexts)
            {
                foreach (var partText in partTexts)
                {
                    var transferGate = GetPdfFields.GetGate(partText);
                    if (transferGate.Contains("TUR"))
                    {
                        Informations inf = new Informations
                        {
                            EmployeeName = GetPdfFields.GetName(partText),
                            TranferGate = transferGate,
                            CardNo = GetPdfFields.GetCardNo(partText),
                            Date = GetPdfFields.GetDate(partText)
                        };

                        listOfInformations.Add(inf);
                    }
                }
            }

        public static List<XmlModel> ParseToXmlModel(List<Informations> employeeList)
        {
            List<XmlModel> xmlModels = new List<XmlModel>();
            var groupByCardNo = employeeList.GroupBy(x => x.CardNo);
            DateTime currentDate = employeeList.FirstOrDefault().Date;
            int days = DateTime.DaysInMonth(currentDate.Year, currentDate.Month);

            for (int day = 1; day <= days; day++)
            {
                DateTime today = new DateTime(currentDate.Year, currentDate.Month, day);
                if ((today.DayOfWeek == DayOfWeek.Saturday) || (today.DayOfWeek == DayOfWeek.Sunday))
                {
                    continue;
                }
                foreach (var cardNo in groupByCardNo)
                {
                    string date = day + "." + currentDate.Month + "." + currentDate.Year;
                    string startTime = cardNo.FirstOrDefault(x => x.Date.Day == day && x.IsEntryRecord == true) == null
                        ? "-" : cardNo.FirstOrDefault(x => x.Date.Day == day && x.IsEntryRecord == true).Date.ToString("dd.MM.yyyy HH:mm").Substring(11, 5);

                    string endTime = cardNo.LastOrDefault(x => x.Date.Day == day && x.IsEntryRecord == false) == null
                        ? "-" : cardNo.LastOrDefault(x => x.Date.Day == day && x.IsEntryRecord == false).Date.ToString("dd.MM.yyyy HH:mm").Substring(11, 5);

                    XmlModel card = new XmlModel
                    {
                        CardNo = cardNo.Key,
                        Date = date,
                        StartTime = startTime,
                        EndTime = endTime,
                        EmployeeName = cardNo.First(x => !x.EmployeeName.Equals(string.Empty)).EmployeeName
                    };
                    xmlModels.Add(card);
                }
            }
            return xmlModels;
        }

        public static void ListToExcel(List<XmlModel> _models, string fileSavePath)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                excel.Workbook.Worksheets.Add("Worksheet2");


                // Target a worksheet
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];

                // Popular header row data
                worksheet.Cells[2, 1].LoadFromCollection(_models);

                FileInfo excelFile = new FileInfo(fileSavePath);
                excel.SaveAs(excelFile);

            }
        }
    }
}
