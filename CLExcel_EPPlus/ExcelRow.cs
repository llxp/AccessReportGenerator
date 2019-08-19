using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CLExcel_EPPlus {
    [ComVisible(true)]
    public class ExcelRow : IEnumerable<ExcelCell>, IEnumerable {
        private readonly List<ExcelCell> m_innerCells = new List<ExcelCell>();
        private readonly Dictionary<string, int> m_columnNames = new Dictionary<string, int>();
        private int m_row = -1;
        public int Row {
            get => m_row;
            set {
                m_row = value;
                foreach(ExcelCell cell in m_innerCells) {
                    cell.Row = m_row;
                }
            }
        }
        public string RowName { get; set; }

        public int Count => m_innerCells.Count;
        public ExcelRow() { }
        public ExcelRow(List<ExcelCell> cells) {
            m_innerCells = cells;
            int previousColumn = 0;
            foreach(ExcelCell cell in m_innerCells) {
                if(cell.Column == -1) {
                    cell.Column = previousColumn;
                    ++previousColumn;
                }
            }
        }

        public ExcelCell this[string columnName] {
            get => GetCell(columnName);
            set => SetCell(columnName, value);
        }

        public ExcelCell this[int column] {
            get => GetCell(column);
            set => SetCell(column, value);
        }

        public ExcelCell GetCell(int column) {
            return m_innerCells.Count > column ? m_innerCells[column] : null;
        }

        public void SetCell(int column, ExcelCell cell) {
            if (m_innerCells.Count > column) {
                m_innerCells[column] = cell;
            }
        }

        public void RemoveAt(int column) {
            if (m_innerCells.Count > column) {
                m_innerCells.RemoveAt(column);
            }
        }

        public ExcelCell GetCell(string columnName) {
            return m_columnNames.ContainsKey(columnName) && m_innerCells.Count > m_columnNames[columnName] ? m_innerCells[m_columnNames[columnName]] : null;
        }

        public void SetCell(string columnName, ExcelCell cell) {
            if (m_columnNames.ContainsKey(columnName) && m_innerCells.Count > m_columnNames[columnName]) {
                m_innerCells[m_columnNames[columnName]] = cell;
            }
        }

        public void SetColumnName(string columnName, int column) {
            if(m_columnNames.ContainsKey(columnName)) {
                m_columnNames[columnName] = column;
            } else {
                m_columnNames.Add(columnName, column);
            }
        }

        public bool FindAddress(string address) {
            for (int i = 0; i < m_innerCells.Count; i++) {
                if (m_innerCells[i].Address == address) {
                    return true;
                }
            }
            return false;
        }

        public void Add(ExcelCell cell) {
            m_innerCells.Add(cell);
            if (cell.Row <= 0) {
                cell.Row = Row;
            }
            if(cell.Column <= 0) {
                cell.Column = m_innerCells.Count > 0 ? m_innerCells.Count - 1 : 0;
            }
        }
        public void Clear() {
            m_innerCells.Clear();
        }

        public IEnumerator<ExcelCell> GetEnumerator() {
            return ((IEnumerable<ExcelCell>)m_innerCells).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator() {
            return ((IEnumerable<ExcelCell>)m_innerCells).GetEnumerator();
        }
    }
}