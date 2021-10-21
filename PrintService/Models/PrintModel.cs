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
        public string format { get; set; }
        public int copies { get; set; }
    }
}
