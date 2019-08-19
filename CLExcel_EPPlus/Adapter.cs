using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using System.Xml;
using OfficeOpenXml;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace CLExcel_EPPlus {
    [ComVisible(false)]
    public class Adapter {
        private ExcelWorksheet m_currentWorkSheet;

        public ExcelRange GetExcelCells() {
            try {
                return m_currentWorkSheet.Cells;
            } catch {
                return null;
            }
        }
        public string FilePath { get; set; } = "";
        public string CurrentWorksheet { get; set; } = "";
        private ExcelPackage pck;
        public Adapter() { }
        public Adapter(string path) {
            FilePath = path;
        }
        public Adapter(string path, string worksheet) {
            bool documentOpened = false;
            while (!documentOpened) {
                try {
                    FilePath = path;
                    CurrentWorksheet = worksheet;
                    FileInfo existingFile = new FileInfo(FilePath);
                    ExcelPackage pck = new ExcelPackage(existingFile);
                    if (pck != null) {
                        this.pck = pck;
                    }
                    if (worksheetExists(worksheet, pck.Workbook.Worksheets)) {
                        ExcelWorksheet worksheet1 = pck.Workbook.Worksheets[worksheet];
                        m_currentWorkSheet = worksheet1;
                    }
                    documentOpened = true;
                } catch {
                    System.Windows.Forms.MessageBox.Show(
                        @"Das Dokument it zur Zeit geöffnet und kann daher nicht geöffnet werden.\n
                    Bitte schließen Sie das Dokument und klicken Sie auf OK.\n
                    Wenn das Programm gestartet ist, können Sie das Dokument gerne wieder öffnen.
                    (Sofern Sie keine weiteren Warnungen sehen)");
                    System.Threading.Thread.Sleep(1000);
                }
            }
        }

        private bool worksheetExists(string worksheetName, ExcelWorksheets worksheets) {
            foreach (ExcelWorksheet worksheet in worksheets) {
                if (worksheet.Name == worksheetName) {
                    return true;
                }
            }
            return false;
        }

        public ExcelRow GetHeaderColumns(ExcelRange row) {
            ExcelRow tempRow = new ExcelRow();
            foreach (ExcelRangeBase firstRowCell in row) {
                ExcelCell cell = new ExcelCell(
                    firstRowCell.Address,
                    firstRowCell.Text,
                    "#" + firstRowCell.Style.Fill.BackgroundColor.Rgb);
                tempRow.Add(cell);
            }
            return tempRow;
        }
        public List<string> GetHeaderColumns(int start = 0) {
            List<string> columnNames = new List<string>();
            foreach (ExcelRangeBase firstRowCell in m_currentWorkSheet.Cells[
                m_currentWorkSheet.Dimension.Start.Row + start,
                m_currentWorkSheet.Dimension.Start.Column,
                m_currentWorkSheet.Dimension.Start.Row + start,
                m_currentWorkSheet.Dimension.End.Column]) {
                columnNames.Add(firstRowCell.Text);
            }
            return columnNames;
        }

        public int GetRowRange() {
            if (m_currentWorkSheet != null && m_currentWorkSheet.Dimension != null) {
                int startRows = m_currentWorkSheet.Dimension.Start.Row;
                int endRows = m_currentWorkSheet.Dimension.End.Row;
                return endRows - startRows;
            }
            return 0;
        }

        public int GetColumnRange() {
            if (m_currentWorkSheet != null && m_currentWorkSheet.Dimension != null) {
                int startCols = m_currentWorkSheet.Dimension.Start.Column;
                int endCols = m_currentWorkSheet.Dimension.End.Column;
                return endCols - startCols;
            }
            return 0;
        }

        public ExcelTable ReadFullTable(int startRow = 0, int startColumn = 0, int endRow = -1, int endColumn = -1) {
            ExcelTable table = new ExcelTable();
            ExcelRange range = GetExcelCells();
            if (range == null) {
                return null;
            }
            int rows = GetRowRange();
            int columns = GetColumnRange();
            if (m_currentWorkSheet.Dimension != null) {
                ExcelCellAddress start = m_currentWorkSheet.Dimension.Start;
                ExcelCellAddress end = m_currentWorkSheet.Dimension.End;
                int endrow = 0;
                int endcolumn = 0;
                endrow = endRow == -1 ? end.Row : start.Row + endRow;
                endcolumn = endColumn == -1 ? end.Column :
                    (endColumn < 0 ? end.Column + endColumn : start.Column + endColumn);
                for (int row = start.Row + startRow; row <= endrow; ++row) { // Row by row...
                    ExcelRow tempRow = new ExcelRow { Row = row };
                    for (int col = start.Column + startColumn; col <= endcolumn; col++) { // ... Cell by cell...
                        string cellText = m_currentWorkSheet.Cells[row, col].Text; // This got me the actual value I needed.
                        string cellColor = m_currentWorkSheet.Cells[row, col].Style.Fill.BackgroundColor.Rgb;
                        string cellAddress = m_currentWorkSheet.Cells[row, col].Address;
                        ExcelCell tempCell = new ExcelCell {
                            Address = cellAddress,
                            Column = col,
                            Row = row,
                            Value = cellText
                        };
                        if (cellColor != null && cellColor.Length > 0) {
                            tempCell.Color = "#" + cellColor;
                        }

                        switch (cellText.ToLowerInvariant()) {
                            case "green":
                            case "grün":
                                tempCell.Color = "#00FF00";
                                tempCell.Value = "";
                                break;
                            case "red":
                            case "Rot":
                                tempCell.Color = "#FF0000";
                                tempCell.Value = "";
                                break;
                            case "blue":
                            case "blau":
                                tempCell.Color = "#0000FF";
                                tempCell.Value = "";
                                break;
                            case "yellow":
                            case "gelb":
                                tempCell.Color = "#FFFF00";
                                tempCell.Value = "";
                                break;
                            case "violet":
                            case "violett":
                            case "pink":
                            case "Rosa":
                            case "lilac":
                            case "lila":
                            case "magenta":
                                tempCell.Color = "#FF00FF";
                                tempCell.Value = "";
                                break;
                            case "cyan":
                            case "türkis":
                            case "aqua":
                                tempCell.Color = "#00FFFF";
                                tempCell.Value = "";
                                break;
                            default:
                                break;

                        }
                        tempRow.Add(tempCell);
                    }
                    table.Add(tempRow);
                }
                if (startRow - 1 < 0) {
                    startRow = 1;
                }
                table.Headers =
                    GetHeaderColumns(
                        m_currentWorkSheet.Cells[
                            start.Row + (startRow - 1),
                            start.Column,
                            1 + (startRow - 1),
                            endcolumn]);
            }
            return table;
        }

        internal static bool sheetExist(ExcelWorkbook workbook, string sheetName) {
                return workbook.Worksheets.Any(sheet => sheet.Name == sheetName);
        }

        public void SaveAs(ExcelDocument document, string saveFilePath) {
            ExcelPackage pck = new ExcelPackage();
            ExcelWorkbook wb = pck.Workbook;
            foreach (KeyValuePair<string, ExcelTable> kvPairTable in document.WorkSheets) {
                ExcelWorksheet worksheet = wb.Worksheets.Add(kvPairTable.Key);
                ExcelTable table = kvPairTable.Value;
                foreach (ExcelRow row in table) {
                    foreach (ExcelCell cell in row) {
                        //if (cell.Value.Length > 0) {
                            ExcelRange range = worksheet.Cells[cell.Row + 1, cell.Column + 1];
                            if (!string.IsNullOrEmpty(cell.Value) && cell.Value[0] == '=') {
                                range.Formula = cell.Value;
                            } else {
                                range.Value = cell.Value;
                            }
                            if (!string.IsNullOrEmpty(cell.CellFormat)) {
                                range.Style.Numberformat.Format = cell.CellFormat;
                            }
                            if (!string.IsNullOrEmpty(cell.Color)) {
                                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(cell.Color));
                            }
                            if (!string.IsNullOrEmpty(cell.TextColor)) {
                                range.Style.Font.Color.SetColor(ColorTranslator.FromHtml(cell.TextColor));
                            }
                            if(cell.FontSize > 0) {
                                range.Style.Font.Size = cell.FontSize;
                            }
                        //}
                        worksheet.Column(cell.Column + 1).AutoFit();
                    }
                }
            }
            if (saveFilePath != null) {
                FileInfo info = new FileInfo(saveFilePath);
                pck.SaveAs(info);
            }
            pck.Dispose();
        }

        public void Save(ExcelTable table, string worksheetName, string saveFilePath = null) {
            if (m_currentWorkSheet != null) {
                FileInfo existingFile = new FileInfo(FilePath);
                if (existingFile != null) {
                    ExcelPackage pck = new ExcelPackage(existingFile);
                    ExcelWorkbook wb = pck.Workbook;
                    if(!sheetExist(wb, worksheetName)) {
                        wb.Worksheets.Add(worksheetName);
                    }
                    ExcelWorksheet worksheet = wb.Worksheets[worksheetName];
                    int startRows = m_currentWorkSheet.Dimension.Start.Row;
                    int endRows = m_currentWorkSheet.Dimension.End.Row;
                    int rowRange = endRows - startRows;
                    foreach (ExcelRow row in table) {
                        if (rowRange > row.Row + 1) {
                            foreach (ExcelCell cell in row) {
                                if (rowRange > row.Row + 1) {
                                    if (cell.Value.Length > 0) {
                                        ExcelRange range = wb.Worksheets[worksheetName]
                                            .Cells[
                                            cell.Row,
                                            cell.Column];
                                        if (range.Formula != null && range.Formula.Length > 0) {
                                            if (cell.Value[0] == '=') {
                                                range.Formula = cell.Value;
                                            }
                                        } else {
                                            range.Value = cell.Value;
                                        }
                                        if (!string.IsNullOrEmpty(cell.CellFormat)) {
                                            range.Style.Numberformat.Format = cell.CellFormat;
                                        }
                                        if(!string.IsNullOrEmpty(cell.Color)) {
                                            range.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(cell.Color));
                                        }
                                        if (!string.IsNullOrEmpty(cell.TextColor)) {
                                            range.Style.Font.Color.SetColor(ColorTranslator.FromHtml(cell.TextColor));
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (saveFilePath != null) {
                        FileInfo info = new FileInfo(saveFilePath);
                        pck.SaveAs(info);
                    } else {
                        pck.Save();
                    }
                    pck.Dispose();
                }
            }
        }
        public List<string> GetAllWorksheets(string file = "") {
            List<string> allWorksheets = new List<string>();
            if (pck != null) {
                ExcelWorkbook workBook = pck.Workbook;
                foreach (ExcelWorksheet worksheet in workBook.Worksheets) {
                    allWorksheets.Add(worksheet.Name);
                }
            } else {
                return null;
            }
            return allWorksheets;
        }
    }
}
