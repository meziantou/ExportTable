using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportTable
{
    public class CsvOptions
    {
        public CsvOptions()
        {
            Separator = ",";
            Quote = "\"";
        }

        public string Separator { get; set; }
        public string Quote { get; set; }
    }
}
