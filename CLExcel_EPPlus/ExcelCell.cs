using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CLExcel_EPPlus {
    [ComVisible(true)]
    public class ExcelCell {
        public int Row { get; set; } = -1;
        public int Column { get; set; } = -1;
        public string Address { get; set; }
        public string Value { get; set; }
        public string Color { get; set; }
        public string CellFormat { get; set; }
        public string TextColor { get; set; }
        public int FontSize { get; set; }
        public object Tag { get; set; }
        public ExcelCell() { Value = ""; }
        public ExcelCell(string text) {
            Value = text;
        }
        public ExcelCell(string address, string text, string color) {
            Address = address;
            Value = text;
            Color = color;
        }
        public ExcelCell(string address, string text) {
            Address = address;
            Value = text;
        }
        public ExcelCell(int row, int column, string text, string color = "") {
            Value = text;
            Row = row;
            Column = column;
            Color = color;
        }
    }
}