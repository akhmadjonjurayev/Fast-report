using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Fast_Report_Web_Api_First.Model
{
    public class Reports
    {
        // Report ID
        public int Id { get; set; }
        // Report File Name
        public string ReportName { get; set; }
    }
    public class ReportQuery
    {
        public string Format { get; set; }
        public bool Inline { get; set; }
        public string Parameter { get; set; }
    }
}
