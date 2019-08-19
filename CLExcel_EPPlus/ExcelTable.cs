using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CLExcel_EPPlus {
    [ComVisible(true)]
    public class ExcelTable : IEnumerable<ExcelRow>, IEnumerable {
        public ExcelRow Headers { get; set; } = new ExcelRow();
        public int Count => m_innerRows != null ? m_innerRows.Count : 0;
        private List<ExcelRow> m_innerRows = new List<ExcelRow>();
        public ExcelTable() { }
        public string TableName { get; set; }

        public ExcelRow this[string rowName] {
            get => m_innerRows.Any(o => o.RowName == rowName) ? m_innerRows.First(o => o.RowName == rowName) : null;
            set {
                if (m_innerRows.Any(o => o.RowName == rowName)) {
                    for (int i = 0; i < m_innerRows.Count; ++i) {
                        if (m_innerRows[i].RowName == rowName) {
                            m_innerRows[i] = value;
                        }
                    }
                }
            }
        }

        public ExcelRow this[int row] {
            get => m_innerRows.Count > row ? m_innerRows[row] : null;
            set {
                if (m_innerRows.Count > row) {
                    m_innerRows[row] = value;
                }
            }
        }

        public ExcelCell this[int row, int column] {
            get {
                if (m_innerRows.Count > row) {
                    ExcelRow tempRow = m_innerRows[row];
                    if (tempRow.Count > column) {
                        ExcelCell tempCell = tempRow[column];
                        return tempCell;
                    } else {
                        return null;
                    }
                } else {
                    return null;
                }
            }
            set {
                if (m_innerRows.Count > row) {
                    ExcelRow tempRow = m_innerRows[row];
                    if (tempRow.Count < column) {
                        ExcelCell tempCell = tempRow[column];
                        tempCell = value;
                    }
                }
            }
        }

        public ExcelRow GetRow(int row) {
            return this[row];
        }

        public ExcelCell GetCell(int row, int column) {
            return this[row, column];
        }

        public void Add(ExcelRow row) {
            row.Row = m_innerRows.Count;
            m_innerRows.Add(row);
            updateRowNumbers();
        }

        public void Insert(ExcelRow row, int rowNumber) {
            if (m_innerRows.Count > rowNumber) {
                m_innerRows.Insert(rowNumber, row);
                updateRowNumbers();
            }
        }

        public void Insert(ExcelRow row, string rowName) {
            if (this[rowName] != null) {
                ExcelRow tempCell = this[rowName];
                m_innerRows.Insert(tempCell.Row, row);
                updateRowNumbers();
            }
        }

        private void updateRowNumbers() {
            int previousRow = 0;
            foreach (ExcelRow rows in m_innerRows) {
                if (rows.Row == -1) {
                    rows.Row = previousRow;
                    ++previousRow;
                }
            }
        }

        public void ForceUpdateRowNumbers() {
            int previousRow = 0;
            foreach (ExcelRow rows in m_innerRows) {
                rows.Row = previousRow;
                ++previousRow;
            }
        }

        public void SetRowName(string rowName, int index) {
            if (m_innerRows.Count > index) {
                m_innerRows[index].RowName = rowName;
            }
        }

        public void Add(string rowName, ExcelRow row) {
            Add(row);
            SetRowName(rowName, m_innerRows.Count - 1);
        }

        public void Clear() {
            m_innerRows.Clear();
        }

        public IEnumerator<ExcelRow> GetEnumerator() {
            return ((IEnumerable<ExcelRow>)m_innerRows).GetEnumerator();
        }

        public void RemoveAt(int index) {
            m_innerRows.RemoveAt(index);
        }

        IEnumerator IEnumerable.GetEnumerator() {
            return ((IEnumerable<ExcelRow>)m_innerRows).GetEnumerator();
        }
    }
}