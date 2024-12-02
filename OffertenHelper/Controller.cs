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
        public static string GetExcelData(string filename, List<string> mappingValues)
        {
            Helper.GetCellValuesFromExcel(filename, mappingValues);

            return "Test";
        }
    }
}
