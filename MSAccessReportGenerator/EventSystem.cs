using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace AccessReportGenerator {
    [ComVisible(true)]
    public class EventSystem {
        private Dictionary<int, string> m_strings = new Dictionary<int, string>();

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate void ExecuteSQLDelegate(int index);
        private event ExecuteSQLDelegate ExecuteSQL;
        public delegate void DataReadyEventHandler(int index);
        protected event DataReadyEventHandler DataReady;

        public void SetDataReady(int index) {
            DataReady?.Invoke(index);
        }

        public void RegisterCallback(long ptr) {
            ExecuteSQL += Marshal.GetDelegateForFunctionPointer<ExecuteSQLDelegate>(new IntPtr(ptr));
        }

        public int GetStringLength(int index) {
            return m_strings[index].Length;
        }

        public int GetStringPartCount(int index) {
            List<string> strings = new List<string>(m_strings[index].SplitBy(250));
            return strings.Count;
        }

        public string GetStringPart(int index, int partIndex) {
            List<string> strings = new List<string>(m_strings[index].SplitBy(250));
            return strings[partIndex];
        }

        public string GetString(int index) {
            return m_strings[index];
        }

        private void setString(int index, string value) {
            if (m_strings.ContainsKey(index)) {
                m_strings[index] = value;
            } else {
                m_strings.Add(index, value);
            }
        }

        public int RegisterString(string value) {
            int newInt = m_strings.Count > 0 ? m_strings.Count : 0;
            setString(newInt, value);
            return newInt;
        }

        protected void onExecuteSQL(int newInt) {
            ExecuteSQL?.Invoke(newInt);
        }
    }

    public static class EnumerableEx {
        public static IEnumerable<string> SplitBy(this string str, int chunkLength) {
            if (string.IsNullOrEmpty(str)) {
                throw new ArgumentException();
            }

            if (chunkLength < 1) {
                throw new ArgumentException();
            }

            for (int i = 0; i < str.Length; i += chunkLength) {
                if (chunkLength + i > str.Length) {
                    chunkLength = str.Length - i;
                }

                yield return str.Substring(i, chunkLength);
            }
        }
    }
}