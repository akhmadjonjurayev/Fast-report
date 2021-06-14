using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Fast_Report_Web_Api_First.Model
{
    public class Person
    {
        public int Id { get; set; }
        public string firstName { get; set; }
        public string lastName { get; set; }
        public DateTime birthday { get; set; }
        public string address { get; set; }
        public string phone { get; set; }
        public byte[] picture { get; set; }
        public string QrCode { get; set; }
        public string html { get; set; }
    }
}
