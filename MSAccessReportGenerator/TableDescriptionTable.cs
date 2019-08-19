using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace AccessReportGenerator {
    [ComVisible(false)]
    public class TableDescriptionTable : List<TableDescriptionTableEntry> {
    }

    [ComVisible(false)]
    public class TableDescriptionTableEntry {
        public int ID { get;set; }
        public string Excel_Row_Header { get; set; }    // used only on access/vba side (cosmetical purpose only)
        public string Access_Column_Name { get; set; }  // used as a reference for either fetching the data from ReportData or requested data in postprocessing
        public string Access_Report_Name { get; set; }  // used only on access/vba side
        //public string Access_Table_Name { get; set; }   // used only on access/vba side
        public int Row_Number { get; set; }  // used for positioning of the cell
        public int Column_Number { get; set; }  // used for positioning of the cell
        public bool Is_Static_Text { get; set; }  // used to indicate if the current cell data should be looked up in the reportData or if a static text should be displayed
        public string Static_Text { get; set; }  // used to statically display text in the cell
        public string Formula { get; set; }  // can be used alternatively if no static text is displayed
        public string SQL_Source_Query { get; set; }  // used for the postprocessing step to dynamically request data from access
        public string SQL_Report_Reference_Fields { get; set; }  // Comma Separated List of Reference Fields from ReportData to replace the {0}, {1}, ... in SQL_Source_Query
        // Cell Attributes
        // --------------------------------
        public string Background_Color { get; set; }
        public string Foreground_Color { get; set; }
        public string Excel_Cell_Format { get; set; }
        public int Font_Size { get; set; }
        // --------------------------------
        public void FromDictionary(Dictionary<string, string> dictionary) {
            foreach (KeyValuePair<string, string> kvPair in dictionary) {
                PropertyInfo property = GetType().GetProperty(kvPair.Key);
                if (property != null && !string.IsNullOrEmpty(kvPair.Value) && kvPair.Value != "NULL") {
                    if (property.PropertyType == typeof(int)) {
                        property.SetValue(this, int.Parse(kvPair.Value));
                    } else if (property.PropertyType == typeof(bool)) {
                        try {
                            property.SetValue(this, bool.Parse(kvPair.Value));
                        } catch {
                            property.SetValue(this, kvPair.Value == "Yes" || kvPair.Value == "Ja" || kvPair.Value == "Wahr" || kvPair.Value == "True");
                        }
                    } else {
                        property.SetValue(this, kvPair.Value);
                    }
                }
            }
        }
    }
}