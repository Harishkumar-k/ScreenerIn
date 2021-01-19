using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScreenerIn_Automation
{
    public class CompanyModel
    {
        public int id { get; set; }
        public string name { get; set; }
        public string url { get; set; }
    }

    public class ExcelModel
    {
        public List<string> comapny { get; set; } = new List<string>();
        public List<string> url { get; set; } = new List<string>();
    }
}
