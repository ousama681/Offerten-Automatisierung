using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelEditor
{
    public class Helper
    {
        public class Mapping
        {
            public string CellReference;
            public string TxtField;

            public Mapping(string txtField, string cellReference)
            {
                this.TxtField = txtField;
                this.CellReference = cellReference;
            }

            public override bool Equals(object? obj)
            {
                Mapping thatObj;
                if (obj is Mapping)
                {
                    thatObj = (Mapping)obj;
                }
                else
                {
                    return false;
                }

                return CellReference.Equals(thatObj.CellReference);
            }
        }



        public static string ReadExcelSheetDefinedCells(string fname, List<Mapping> definedCells, bool firstRowIsHeader = true)
        {
            List<string> Headers = new List<string>();
            DataTable dt = new DataTable();
            string output = "";

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fname, false))
            {
                Sheet sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
                int counter = 0;

                IEnumerable<Cell> cells = worksheet.GetFirstChild<SheetData>().Descendants<Cell>();

                foreach (Cell cell in cells)
                {
                    if (definedCells.Contains(new Mapping("", cell.CellReference)))
                    {
                        string txtField = definedCells.Where(map => map.CellReference.Equals(cell.CellReference)).First().TxtField;

                        output += cell.CellReference + ": " + GetCellValue(doc, cell) + " ==> " + txtField + "\r\n";
                    }
                }

            }
            return output;
        }

        private static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            string value = cell.CellValue == null ? "" : cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.ElementAt(int.Parse(value)).InnerText;
            }
            return value;
        }
    }
}
