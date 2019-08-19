using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessReportGenerator {
    public class RowHeaderContainer {
        private Dictionary<int, string> m_anonymousRowsTranslationTable = new Dictionary<int, string>();
        private Dictionary<string, int> m_anonymousRowsReverseTranslationTable = new Dictionary<string, int>();

        public string GetRowHeader(int rowIndex) {
            string randomString = generateRandomString();
            if (!m_anonymousRowsTranslationTable.ContainsKey(rowIndex)) {
                m_anonymousRowsTranslationTable.Add(rowIndex, randomString);
                m_anonymousRowsReverseTranslationTable.Add(randomString, rowIndex);
                return randomString;
            } else {
                return m_anonymousRowsTranslationTable[rowIndex];
            }
        }

        public string LookupRowHeader(int rowIndex) {
            return m_anonymousRowsTranslationTable.ContainsKey(rowIndex) ? m_anonymousRowsTranslationTable[rowIndex] : "";
        }

        private static Random random = new Random();
        private string generateRandomString(int length = 20) {
            const string CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            string randomString = new string(Enumerable.Repeat(CHARS, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
            while (m_anonymousRowsReverseTranslationTable.ContainsKey(randomString)) {
                randomString = new string(Enumerable.Repeat(CHARS, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
            }
            return randomString;
        }
    }
}
