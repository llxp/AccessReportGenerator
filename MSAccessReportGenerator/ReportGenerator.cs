using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.EnterpriseServices;
using CLExcel_EPPlus;
using System.Text.RegularExpressions;
using System.Collections;

namespace AccessReportGenerator {
    [ComVisible(true)]
    public class ReportGenerator : EventSystem {

        private Dictionary<string, TableDescriptionTable> m_tableDescriptionTable = new Dictionary<string, TableDescriptionTable>();
        private Dictionary<string, string> m_reportData = new Dictionary<string, string>();

        private Dictionary<int, int> m_postprocessingData = new Dictionary<int, int>();
        private Dictionary<int, int> m_postprocessingReverseData = new Dictionary<int, int>();
        private List<int> m_dataReadyList = new List<int>();
        private Dictionary<int, List<Dictionary<string, string>>> m_additionalData = new Dictionary<int, List<Dictionary<string, string>>>();
        private Dictionary<int, string> m_worksheetList = new Dictionary<int, string>();

        private RowHeaderContainer m_rowHeaderContainer = new RowHeaderContainer();

        private ExcelDocument m_reportTable = new ExcelDocument();

        // function will be called by the vba script to add a new row of the TableDescriptionTable
        public void AddDescriptionTableEntry(string worksheetName, Scripting.Dictionary descriptionTableEntry) {
            Dictionary<string, string> descriptionEntry = ConversionFunctions.TranslateDictionary(descriptionTableEntry);
            TableDescriptionTableEntry reportDescriptionData = new TableDescriptionTableEntry();
            reportDescriptionData.FromDictionary(descriptionEntry);
            if (m_tableDescriptionTable.ContainsKey(worksheetName)) {
                m_tableDescriptionTable[worksheetName].Add(reportDescriptionData);
            } else {
                m_tableDescriptionTable.Add(worksheetName, new TableDescriptionTable());
                m_tableDescriptionTable[worksheetName].Add(reportDescriptionData);
            }
        }

        // function will be called by the vba script to set the actual data from the current data record
        public void SetReportData(Scripting.Dictionary reportData) {
            m_reportData = ConversionFunctions.TranslateDictionary(reportData);
        }

        // function will be called from the callback function, when there is requested new data from the report generator
        // report generator needds more data --> call registered callback function in the vba script --> vba callback function calls the function "AddAdditionalData"
        public void AddAdditionalData(Scripting.Dictionary additionalData, int index) {
            if (m_additionalData.ContainsKey(index)) {
                m_additionalData[index].Add(new Dictionary<string, string>(ConversionFunctions.TranslateDictionary(additionalData)));
            } else {
                m_additionalData.Add(index, new List<Dictionary<string, string>>() { new Dictionary<string, string>(ConversionFunctions.TranslateDictionary(additionalData)) });
            }
        }

        // called at the end or the beginning of the vba script to clear all previous configurations
        public void ClearDescriptionTableData() {
            m_tableDescriptionTable.Clear();
        }

        // function will be called by the vba script
        // to start the actual report generation process
        // of one excel worksheet --> specified by the tableName parameter
        public void GenerateLocationReport(string fileName, string tableName) {
            m_reportTable.FileName = fileName;
            if (m_reportTable.WorkSheets.ContainsKey(tableName)) {
                m_reportTable.WorkSheets[tableName] = new ExcelTable();
            } else {
                m_reportTable.WorkSheets.Add(tableName, new ExcelTable());
            }
            // prepare the template according to the current configuration
            // and prefill the template with the data from the current dataset
            populateExcelTable(tableName);
            DataReady += reportGenerator_DataReady;
            // check if there is placeholder data in the template
            // to be filled in with additional data from the access database
            requestPostprocessingData(tableName);
        }

        // function will be called at the end of the VBA script
        public void SaveFile() {
            Adapter excelAdapter = new Adapter(m_reportTable.FileName);
            excelAdapter.SaveAs(m_reportTable, m_reportTable.FileName);
        }

        // return a sorted version of the tableDescriptionTable filtered by the specified worksheet name
        private IOrderedEnumerable<TableDescriptionTableEntry> getSortedDescriptionTable(string worksheetName) {
            return m_tableDescriptionTable[worksheetName].OrderBy(o => o.Row_Number).ThenBy(o => o.Column_Number);
        }

        // called by the function "GenerateLocationReport"
        // prepares the template according to the current configuration in the tableDescriptionTable
        // and prefills the template with the data from the currently selected dataset
        private void populateExcelTable(string worksheetName) {
            foreach (TableDescriptionTableEntry descriptionEntry in getSortedDescriptionTable(worksheetName)) {
                string originalRowHeader = descriptionEntry.Excel_Row_Header;
                int rowNumber = descriptionEntry.Row_Number;
                string rowHeader = m_rowHeaderContainer.GetRowHeader(rowNumber);

                if (descriptionEntry.Is_Static_Text) {
                    // add or set a static text to the specified cell
                    string staticText =
                        descriptionEntry.Static_Text.NotEmpty() ?
                            descriptionEntry.Static_Text :
                            descriptionEntry.Formula.NotEmpty() ?
                                descriptionEntry.Formula : "";
                    addOrSetCells(worksheetName, rowHeader, staticText, descriptionEntry);
                } else if (descriptionEntry.Access_Column_Name.NotEmpty()
                    && descriptionEntry.SQL_Source_Query.Empty()) {
                    // add or set data from the reportData according
                    // to the description table to the specified cell
                    string staticText = 
                        m_reportData[descriptionEntry.Access_Column_Name].NotEmpty() ?
                            m_reportData[descriptionEntry.Access_Column_Name] : "";
                    if (m_reportData.Count(key => key.Key == descriptionEntry.Access_Column_Name) > 0) {
                        addOrSetCells(worksheetName, rowHeader, staticText, descriptionEntry);
                    }
                } else if (descriptionEntry.Access_Column_Name.NotEmpty()
                    && descriptionEntry.SQL_Source_Query.NotEmpty()) {
                    // add a placeholder cell, which is marked with "DEL"
                    // to identify and delete it later in the postprocessing step
                    string staticText = descriptionEntry.Access_Column_Name;
                    addOrSetCells(worksheetName, rowHeader, staticText, descriptionEntry, "DEL");
                }
            }
        }

        private void addOrSetCells(string worksheetName, string rowHeader, string text, TableDescriptionTableEntry descriptionEntry, object tag = null) {
            if (m_reportTable.WorkSheets[worksheetName][rowHeader] != null) {
                addCellToExistingRow(worksheetName, rowHeader, text, descriptionEntry, tag);
            } else {
                addCells(worksheetName, rowHeader, text, descriptionEntry, tag);
            }
        }

        private void addCellToExistingRow(string worksheetName, string rowHeader, string text, TableDescriptionTableEntry descriptionEntry, object tag = null) {
            m_reportTable.WorkSheets[worksheetName][rowHeader].Add(new ExcelCell(text) {
                Color = descriptionEntry.Background_Color.NotEmpty() ? descriptionEntry.Background_Color : "",
                CellFormat = descriptionEntry.Excel_Cell_Format.NotEmpty() ? descriptionEntry.Excel_Cell_Format : "",
                TextColor = descriptionEntry.Foreground_Color.NotEmpty() ? descriptionEntry.Foreground_Color : "",
                Tag = tag,
                Column = descriptionEntry.Column_Number,
                Row = descriptionEntry.Row_Number,
                FontSize = descriptionEntry.Font_Size
            });
        }

        private void addCells(string worksheetName, string rowHeader, string text, TableDescriptionTableEntry descriptionEntry, object tag = null) {
            m_reportTable.WorkSheets[worksheetName].Add(rowHeader, new ExcelRow(new List<ExcelCell>() {
                new ExcelCell(text) {
                    Color = descriptionEntry.Background_Color.NotEmpty() ? descriptionEntry.Background_Color : "",
                    CellFormat = descriptionEntry.Excel_Cell_Format.NotEmpty() ? descriptionEntry.Excel_Cell_Format : "",
                    TextColor = descriptionEntry.Foreground_Color.NotEmpty() ? descriptionEntry.Foreground_Color : "",
                    Tag = tag,
                    Column = descriptionEntry.Column_Number,
                    Row = descriptionEntry.Row_Number,
                    FontSize = descriptionEntry.Font_Size
                } }));
        }

        // called by the function "GenerateLocationReport"
        // when the template has already been prefilled out by the function "populateExcelTable"
        // replaces the previously placed placeholder cells
        // with actual data by requesting more data from the access database
        // using the registered callback function
        private void requestPostprocessingData(string worksheetName) {
            m_postprocessingData.Clear();
            m_postprocessingReverseData.Clear();
            m_dataReadyList.Clear();
            m_additionalData.Clear();
            List<int> registeredIndizes = new List<int>();
            foreach (TableDescriptionTableEntry descriptionEntry in getSortedDescriptionTable(worksheetName)) {
                if (descriptionEntry.SQL_Source_Query.NotEmpty()) {
                    string sqlSourceQuery = descriptionEntry.SQL_Source_Query;
                    // replace the placeholders in the query with the actual data from the "parent" report
                    sqlSourceQuery = replacePlaceholders(sqlSourceQuery, descriptionEntry.SQL_Report_Reference_Fields);
                    // if the replacement was not successful don't proceed
                    if (sqlSourceQuery.NotEmpty()) {
                        // register the string and get a registration index back
                        int registeredIndex = RegisterString(sqlSourceQuery);
                        // register the name of the current worksheet internally using the registration index
                        if(registeredIndex == 8) {
                            // specific breakpoint for debugging
                            Console.WriteLine("");
                        }
                        // check if the current worksheet name is already registered using the index.
                        // if it is already registered overwrite the worksheet name
                        if (m_worksheetList.ContainsKey(registeredIndex)) {
                            m_worksheetList[registeredIndex] = worksheetName;
                        } else {
                            m_worksheetList.Add(registeredIndex, worksheetName);
                        }
                        // register the TableDescriptionTableEntry ID using the registration index for the postprocessing step
                        // to check later to see if all the data requested has been send to the c# com dll from access
                        // and only start then with the postprocessing step, when the data has completely send to the c# dll
                        setPostprocessingData(registeredIndex, descriptionEntry.ID);
                        registeredIndizes.Add(registeredIndex);
                    }
                    //m_postprocessingData.Add(m_rowHeaderContainer.LookupRowHeader(descriptionEntry.Row_Number));
                }
            }
            foreach(int registeredIndex in registeredIndizes) {
                onExecuteSQL(registeredIndex);
            }
        }

        private string replacePlaceholders(string sqlSourceQuery, string referenceFields) {
            if(Regex.IsMatch(sqlSourceQuery, @"{[-|+]*\d+}")) {
                string[] splittedReferenceFields = referenceFields.Split(',');
                List<string> cleanedSplittedReferenceFields = new List<string>();
                foreach(string referenceField in splittedReferenceFields) {
                    string referenceFieldTemp = referenceField.Trim();
                    if(m_reportData.ContainsKey(referenceFieldTemp)) {
                        cleanedSplittedReferenceFields.Add(m_reportData[referenceFieldTemp]);
                    }
                }
                if(GetDistinctRegexCount(sqlSourceQuery, @"{[-|+]*\d+}") == cleanedSplittedReferenceFields.Count) {
                    sqlSourceQuery = string.Format(sqlSourceQuery, cleanedSplittedReferenceFields.ToArray());
                    return sqlSourceQuery;
                } else {
                    return "";
                }
            }
            return sqlSourceQuery;
        }

        public string Trim(string inputString) {
            return inputString.Trim();
        }

        public string Format(string inputString, ArrayList parameter) {
            string str = string.Format(inputString, ConversionFunctions.TranslateArrayList(parameter).ToArray());
            return str;
        }

        public int GetDistinctRegexCount(string input, string pattern) {
            MatchCollection matches = Regex.Matches(input, pattern);
            List<string> distinctMatches = new List<string>();
            foreach(Match match in matches) {
                if(match.Success) {
                    if(!distinctMatches.Contains(match.Value)) {
                        distinctMatches.Add(match.Value);
                    }
                }
            }
            return distinctMatches.Count;
        }

        private void setPostprocessingData(int registeredIndex, int id) {
            if (m_postprocessingData.ContainsKey(registeredIndex)) {
                m_postprocessingData[registeredIndex] = id;
                m_postprocessingReverseData[id] = registeredIndex;
            } else {
                m_postprocessingData.Add(registeredIndex, id);
                m_postprocessingReverseData.Add(id, registeredIndex);
            }
        }

        private void reportGenerator_DataReady(int index) {
            if (!m_dataReadyList.Contains(index)) {
                m_dataReadyList.Add(index);
                if (m_dataReadyList.Count == m_additionalData.Count
                    && m_dataReadyList.Count == m_postprocessingData.Count) {
                    //postprocessAdditionalData(m_worksheetList[index], index);
                    postProcessData(m_worksheetList[index]);
                    cleanTable(m_worksheetList[index]);
                }
            }
        }

        /*private static string[] getAddonFunctions() {
            List<string> methods = new List<string>();
            foreach(System.Reflection.MethodInfo method in typeof(SQLAddonFunctions).GetMethods()) {
                if(method != null) {
                    methods.Add(method.Name);
                }
            }
            return methods.ToArray();
        }

        private void postprocessAdditionalData(string worksheetName, int index) {
            string query = GetString(index);
            string[] addonFunctions = getAddonFunctions();
            foreach(string function in addonFunctions) {
                if(query.Contains(function)) {
                    int pos = query.IndexOf(function);
                    if(pos != -1) {

                    }
                }
            }
        }*/

        private List<Dictionary<string, string>> getDataset(int registeredIndex) {
            return m_additionalData[registeredIndex];
        }

        private void postProcessData(string worksheetName) {
            foreach (TableDescriptionTableEntry descriptionEntry in getSortedDescriptionTable(worksheetName)) {
                //if (descriptionEntry.Access_Column_Name != null && descriptionEntry.Access_Column_Name.Contains("V_Dash") && worksheetName == "V dash & Security 101") {
                    // only for debugging purposes
                    //Console.WriteLine("test...");
                //}

                if (descriptionEntry.SQL_Source_Query.NotEmpty()) {

                    // lookup using the row header name because the table
                    // will be increased during the postprocessing step
                    // and the row numbers are not matching with
                    // the row number from the description table anymore
                    string rowHeader = m_rowHeaderContainer.LookupRowHeader(descriptionEntry.Row_Number);
                    ExcelRow row = m_reportTable.WorkSheets[worksheetName][rowHeader];

                    if (row != null) {

                        // the registered index which was previously registered during requesting postprocessing data step
                        int registeredIndex = m_postprocessingReverseData[descriptionEntry.ID];
                        // dataset to be inserted in the postprocessing step
                        List<Dictionary<string, string>> dataset = getDataset(registeredIndex);
                        // +1 because next row should be targetted for inserting a new cell
                        // the current row is only a placeholder row and was marked with "DEL" to delete it later
                        int rowNumber = row.Row + 1;
                        int rowNumberBegin = rowNumber;

                        foreach (Dictionary<string, string> record in dataset) {
                            //if (descriptionEntry.Access_Column_Name != null && descriptionEntry.Access_Column_Name.Contains("V_Dash") && worksheetName == "V dash & Security 101" && (rowNumber - 3) == 1) {
                                // only for debugging purposes
                                //Console.WriteLine("test...");
                            //}

                            ExcelCell currentCell = row.Where(cell => cell.Column == descriptionEntry.Column_Number).FirstOrDefault();
                            if (currentCell != null &&
                                currentCell.Tag != null &&
                                ((string)currentCell.Tag) == "DEL" &&
                                rowNumber == rowNumberBegin) {
                                // the current cell is the first cell of the column / the first row
                                // the current cell was marked with "DEL" so it is a placeholder cell
                                // the current cell will be overritten and the "DEL" mark will be removed
                                // from the cell
                                currentCell.Tag = null;
                                currentCell.Value = record.Exists(descriptionEntry.Access_Column_Name) ? record[descriptionEntry.Access_Column_Name].NotEmpty() ? record[descriptionEntry.Access_Column_Name] : "" : "";
                            } else {
                                ExcelRow nextRow = m_reportTable.WorkSheets[worksheetName][rowNumber];
                                if (nextRow != null) {
                                    if (nextRow.RowName.NotEmpty()) {
                                        // the next row has already a row name aka random string
                                        // aka placeholder cell marked with "DEL" --> not marked anymore
                                        // but still begin of a new row / new dataset
                                        // therefore it needs to be added a new row at the specified index
                                        if (nextRow.Where(cell => cell.Column == descriptionEntry.Column_Number).FirstOrDefault() == null) {
                                            // the next row had already got a name but the cell in the next row is still available
                                            reusePreviousRow(worksheetName, descriptionEntry, record, rowNumber);
                                        } else {
                                            // the next row had already got a name and the cell is also not available anymore
                                            addNewRowBetween(worksheetName, descriptionEntry, record, rowNumber);
                                        }
                                    } else {
                                        // a writable new row was already created in a previous loop run, so reuse that row
                                        reusePreviousRow(worksheetName, descriptionEntry, record, rowNumber);
                                    }
                                } else {
                                    // the end of the table was reached
                                    // so create a new row to the table
                                    int rowCount = m_reportTable.WorkSheets[worksheetName].Count;
                                    // check if the rowCount of the created table is larger then the last row index of the cell to be created
                                    // check also if the rowCount is larger or equal to the dataset count
                                    // 
                                    if (rowCount > rowNumber - 1 &&
                                        rowCount - row.Row >= dataset.Count) {
                                        reusePreviousRow(worksheetName, descriptionEntry, record, rowNumber);
                                    } else {
                                        addNewRowToTheEnd(worksheetName, descriptionEntry, record, rowNumber);
                                    }
                                }
                            }
                            // increase the row number
                            ++rowNumber;
                        }
                        //m_reportTable.Insert(new ExcelRow(new List<ExcelCell>() { new ExcelCell() }), row.Row);
                    }
                }
            }
        }

        private void addNewRowBetween(string worksheetName, TableDescriptionTableEntry descriptionEntry, Dictionary<string, string> record, int rowNumber) {
            if (descriptionEntry.Column_Number == 0) {
                // create a new row at the specified index with the value cell as the first cell
                m_reportTable.WorkSheets[worksheetName].Insert(
                new ExcelRow(new List<ExcelCell>() {
                    new ExcelCell(
                        record.Exists(descriptionEntry.Access_Column_Name) ? record[descriptionEntry.Access_Column_Name].NotEmpty() ? record[descriptionEntry.Access_Column_Name] : "" : "") {
                        CellFormat = descriptionEntry.Excel_Cell_Format.NotEmpty() ? descriptionEntry.Excel_Cell_Format : "",
                        Color = descriptionEntry.Background_Color.NotEmpty() ? descriptionEntry.Background_Color : "",
                        TextColor = descriptionEntry.Foreground_Color.NotEmpty() ? descriptionEntry.Foreground_Color : "",
                        FontSize = descriptionEntry.Font_Size
                    }
                }), rowNumber);
            } else {
                int blankColumns = descriptionEntry.Column_Number;
                ExcelRow currentRow = new ExcelRow();  // create new blank row
                addBlankCells(ref currentRow, blankColumns);  // add some blank cells to the new row
                currentRow.Add(new ExcelCell(
                    record.Exists(descriptionEntry.Access_Column_Name) ?
                    record[descriptionEntry.Access_Column_Name].NotEmpty() ? record[descriptionEntry.Access_Column_Name] : ""
                    : "") {
                    CellFormat = descriptionEntry.Excel_Cell_Format.NotEmpty() ? descriptionEntry.Excel_Cell_Format : "",
                    Color = descriptionEntry.Background_Color.NotEmpty() ? descriptionEntry.Background_Color : "",
                    TextColor = descriptionEntry.Foreground_Color.NotEmpty() ? descriptionEntry.Foreground_Color : "",
                    FontSize = descriptionEntry.Font_Size
                });
                m_reportTable.WorkSheets[worksheetName].Insert(currentRow, rowNumber);
            }
        }

        private void addNewRowToTheEnd(string worksheetName, TableDescriptionTableEntry descriptionEntry, Dictionary<string, string> record, int rowNumber) {
            if (descriptionEntry.Column_Number == 0) {
                // create a new row  at the specified row index with the first cell as the value cell
                m_reportTable.WorkSheets[worksheetName].Add(
                new ExcelRow(new List<ExcelCell>() {
                    new ExcelCell(
                        record.Exists(descriptionEntry.Access_Column_Name) ? record[descriptionEntry.Access_Column_Name].NotEmpty() ? record[descriptionEntry.Access_Column_Name] : "" : "") {
                        CellFormat = descriptionEntry.Excel_Cell_Format.NotEmpty() ? descriptionEntry.Excel_Cell_Format : "",
                        Color = descriptionEntry.Background_Color.NotEmpty() ? descriptionEntry.Background_Color : "",
                        TextColor = descriptionEntry.Foreground_Color.NotEmpty() ? descriptionEntry.Foreground_Color : "",
                        FontSize = descriptionEntry.Font_Size
                    }
                }));
            } else {
                // create a new row with some blank cells in front of the actual value cell
                int blankColumns = descriptionEntry.Column_Number;
                ExcelRow newRow = new ExcelRow();
                addBlankCells(ref newRow, blankColumns);  // add some blank cells to the new row
                newRow.Add(new ExcelCell(
                    record.Exists(descriptionEntry.Access_Column_Name) ? record[descriptionEntry.Access_Column_Name].NotEmpty() ? record[descriptionEntry.Access_Column_Name] : "" : "") {
                    CellFormat = descriptionEntry.Excel_Cell_Format.NotEmpty() ? descriptionEntry.Excel_Cell_Format : "",
                    Color = descriptionEntry.Background_Color.NotEmpty() ? descriptionEntry.Background_Color : "",
                    TextColor = descriptionEntry.Foreground_Color.NotEmpty() ? descriptionEntry.Foreground_Color : "",
                    FontSize = descriptionEntry.Font_Size
                });
                m_reportTable.WorkSheets[worksheetName].Add(newRow);  // add the new row to the table
            }
        }

        private void reusePreviousRow(string worksheetName, TableDescriptionTableEntry descriptionEntry, Dictionary<string, string> record, int rowNumber) {
            ExcelRow currentRow = m_reportTable.WorkSheets[worksheetName][rowNumber - 1];
            if (descriptionEntry.Column_Number == 0) {
                if (currentRow[0] == null && currentRow.Count == 0) {
                    // the first cell does not exist and therefore needs to be created
                    currentRow.Add(new ExcelCell(
                        record.Exists(descriptionEntry.Access_Column_Name) ? stringNotEmpty(record[descriptionEntry.Access_Column_Name]) ? record[descriptionEntry.Access_Column_Name] : "" : "") {
                        CellFormat = stringNotEmpty(descriptionEntry.Excel_Cell_Format) ? descriptionEntry.Excel_Cell_Format : "",
                        Color = stringNotEmpty(descriptionEntry.Background_Color) ? descriptionEntry.Background_Color : "",
                        TextColor = stringNotEmpty(descriptionEntry.Foreground_Color) ? descriptionEntry.Foreground_Color : "",
                        FontSize = descriptionEntry.Font_Size
                    });
                } else {
                    // the first cell will be replaced
                    currentRow[0].Value =
                        record.Exists(descriptionEntry.Access_Column_Name) ? stringNotEmpty(record[descriptionEntry.Access_Column_Name]) ? record[descriptionEntry.Access_Column_Name] : "" : "";
                    currentRow[0].TextColor = stringNotEmpty(descriptionEntry.Foreground_Color) ? descriptionEntry.Foreground_Color : "";
                    currentRow[0].Color = stringNotEmpty(descriptionEntry.Background_Color) ? descriptionEntry.Background_Color : "";
                    currentRow[0].CellFormat = stringNotEmpty(descriptionEntry.Excel_Cell_Format) ? descriptionEntry.Excel_Cell_Format : "";
                    currentRow[0].FontSize = descriptionEntry.Font_Size;
                }
            } else if(descriptionEntry.Column_Number > 0) {
                if (descriptionEntry.Column_Number > currentRow.Count) {
                    int blankColumns = descriptionEntry.Column_Number - currentRow.Count;
                    // Example: | 123 | 456 | BlankCell | BlankCell | NewValue |
                    // currentRow.Count = 2
                    // blankColumns = descriptionEntry.Column_Number - currentRow.Count = 4 - 2 = 2
                    // blankColumns = 2
                    // Add 2 new blank cells/columns
                    addBlankCells(ref currentRow, blankColumns);

                    currentRow.Add(new ExcelCell(record.Exists(descriptionEntry.Access_Column_Name) ? stringNotEmpty(record[descriptionEntry.Access_Column_Name]) ? record[descriptionEntry.Access_Column_Name] : "" : "") {
                        CellFormat = stringNotEmpty(descriptionEntry.Excel_Cell_Format) ? descriptionEntry.Excel_Cell_Format : "",
                        Color = stringNotEmpty(descriptionEntry.Background_Color) ? descriptionEntry.Background_Color : "",
                        TextColor = stringNotEmpty(descriptionEntry.Foreground_Color) ? descriptionEntry.Foreground_Color : "",
                        FontSize = descriptionEntry.Font_Size
                    });
                } else if (descriptionEntry.Column_Number < currentRow.Count) {
                    // cell somewhere in the middle should be replaced:
                    // index = descriptionEntry.Column_Number
                    // length = currentRow.Count
                    // index < length
                    currentRow[descriptionEntry.Column_Number].Value =
                        record.Exists(descriptionEntry.Access_Column_Name) ? stringNotEmpty(record[descriptionEntry.Access_Column_Name]) ? record[descriptionEntry.Access_Column_Name] : "" : "";
                    currentRow[descriptionEntry.Column_Number].TextColor = stringNotEmpty(descriptionEntry.Foreground_Color) ? descriptionEntry.Foreground_Color : "";
                    currentRow[descriptionEntry.Column_Number].Color = stringNotEmpty(descriptionEntry.Background_Color) ? descriptionEntry.Background_Color : "";
                    currentRow[descriptionEntry.Column_Number].CellFormat = stringNotEmpty(descriptionEntry.Excel_Cell_Format) ? descriptionEntry.Excel_Cell_Format : "";
                    currentRow[descriptionEntry.Column_Number].FontSize = descriptionEntry.Font_Size;
                } else if (currentRow.Count == descriptionEntry.Column_Number) {
                    // there should be added a new cell to the end of the row
                    currentRow.Add(new ExcelCell(
                        record.Exists(descriptionEntry.Access_Column_Name) ? stringNotEmpty(record[descriptionEntry.Access_Column_Name]) ? record[descriptionEntry.Access_Column_Name] : "" : "") {
                        CellFormat = stringNotEmpty(descriptionEntry.Excel_Cell_Format) ? descriptionEntry.Excel_Cell_Format : "",
                        Color = stringNotEmpty(descriptionEntry.Background_Color) ? descriptionEntry.Background_Color : "",
                        TextColor = stringNotEmpty(descriptionEntry.Foreground_Color) ? descriptionEntry.Foreground_Color : "",
                        FontSize = descriptionEntry.Font_Size
                    });
                }
            }
        }

        private void addBlankCells(ref ExcelRow row, int count) {
            for (int i = 0; i < count; ++i) {
                row.Add(new ExcelCell());
            }
        }

        private void cleanTable(string worksheetName) {
            while (true) {
                int i = 0;
                for (; i < m_reportTable.WorkSheets[worksheetName].Count; ++i) {
                    for (int x = 0; x < m_reportTable.WorkSheets[worksheetName][i].Count; ++x) {
                        ExcelCell cell = m_reportTable.WorkSheets[worksheetName][i][x];
                        if (cell.Tag != null) {
                            m_reportTable.WorkSheets[worksheetName][i].RemoveAt(x);
                            if (m_reportTable.WorkSheets[worksheetName][i].Count == 0) {
                                m_reportTable.WorkSheets[worksheetName].RemoveAt(i);
                                m_reportTable.WorkSheets[worksheetName].ForceUpdateRowNumbers();
                                break;
                            }
                        }
                    }
                }
                if (i >= m_reportTable.WorkSheets[worksheetName].Count - 1) {
                    break;
                }
            }
        }

        private static bool stringNotEmpty(string str) {
            return !string.IsNullOrEmpty(str) && str != "NULL";
        }

        private static bool stringEmpty(string str) {
            return string.IsNullOrEmpty(str) || str == "NULL";
        }
    }

    // extension class for convenience functions
    public static class DictionaryExtensionClass {
        public static bool Exists<V>(this Dictionary<string, V> obj, string key) {
            return obj.Any(o => o.Key == key);
        }

        public static bool NotEmpty(this string str) {
            return !string.IsNullOrEmpty(str) && str != "NULL";
        }

        public static bool Empty(this string str) {
            return string.IsNullOrEmpty(str) || str == "NULL";
        }
    }
}