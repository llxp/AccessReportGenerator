using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessReportGenerator {
    class ConversionFunctions {
        public static Dictionary<string, string> TranslateDictionary(Scripting.Dictionary dictionary) {
            Dictionary<string, string> output = new Dictionary<string, string>();
            foreach (object kvPair in dictionary) {
                string key = kvPair.ToString();
                string value = dictionary.get_Item(kvPair.ToString()).ToString();
                output.Add(key, value);
            }
            return output;
        }

        public static List<string> TranslateArrayList(ArrayList parameter) {
            List<string> output = new List<string>();
            foreach(string str in parameter) {
                output.Add(str);
            }
            return output;
        }
    }
}