using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CLExcel_EPPlus {
    public class ExcelDocument {
        public Dictionary<string, ExcelTable> WorkSheets { get; set; } = new Dictionary<string, ExcelTable>();
        public string FileName { get; set; }
    }
}
