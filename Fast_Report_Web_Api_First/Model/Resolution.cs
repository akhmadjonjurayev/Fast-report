using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Fast_Report_Web_Api_First.Model
{
    public class Resolution
    {
        public string Company { get; set; }
        public string Director { get; set; }
        public DateTime DateTimeNow { get; set; }
        public string FullCompanyName { get; set; }
    }
    public class ResolutionPerson
    {
        public string Author { get; set; }
        public string Persons { get; set; }
        public string Message { get; set; }
        public DateTime DeadLine { get; set; }
        public string Control { get; set; }
        public string Director { get; set; }
        public DateTime DateTimeNow { get; set; }
    }
}
