using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PrintService.Models
{
    public class PrintModel
    {
        public string printer { get; set; }
        public string doc { get; set; }
        public string name_emited { get; set; }
        public string report_title { get; set; }
        public string column1 { get; set; }
        public string column2 { get; set; }
        public string column3 { get; set; }
        public string text_info { get; set; }
        public string format { get; set; }
        public int copies { get; set; }
    }
}
