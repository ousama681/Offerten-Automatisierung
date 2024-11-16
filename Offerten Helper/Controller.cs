using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelEditor;
using static ExcelEditor.Helper;

namespace Offerten_Helper
{
    public class Controller
    {
        public static string GetExcelData(string filename, bool isHeaderOn)
        {
            List<Mapping> mappings = new List<Mapping>();

            mappings.Add(new Mapping("RandomTextfield", "B3"));
            mappings.Add(new Mapping("RandomTextfield", "B32"));
            mappings.Add(new Mapping("RandomTextfield", "B33"));

            return Helper.ReadExcelSheetDefinedCells(filename, mappings, isHeaderOn);
        }
    }
}
