using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfParser
{
    public class Informations
    {
        public string EmployeeName { get; set; }
        public string TranferGate { get; set; }
        public string CardNo { get; set; }
        public DateTime Date { get; set; }

        public bool IsEntryRecord
        {
            get
            {
                return TranferGate.Contains("GIRIS");
            }
        }
    }
}
