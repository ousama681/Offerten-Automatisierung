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

        public static Dictionary<string, string> GetCellValuesFromExcel(string fileName, List<string> mappingKeys)
        {
            Dictionary<string, string> mappedValues = new Dictionary<string, string>();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))

                foreach (var key in mappingKeys)
                {
                    Sheet sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                    Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
                    IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
                    IEnumerable<Cell> cells = worksheet.GetFirstChild<SheetData>().Descendants<Cell>();


                    WorkbookPart workbookPart = doc.WorkbookPart;
                    DefinedNames definedNames = workbookPart.Workbook.DefinedNames;

                    if (definedNames == null)
                    {
                        throw new ArgumentNullException("There are no unique names defined for cells in the excel sheet.");
                    }

                    foreach (DefinedName dn in definedNames)
                    {
                        if (dn.Name.Equals(key))
                        {
                            string cellReference = dn.Text;
                            string cellValue = GetCellValueFromDefinedName(workbookPart, dn.Name);
                            mappedValues.Add(cellReference, cellValue);
                        }
                    }
                }


            if (mappedValues.Count == 0)
            {
                throw new ArgumentNullException("There were no matching defined names found in the excel sheet that matched the provided ones from the tool.");
            }

            return mappedValues;
        }

        public static string GetCellValueFromDefinedName(WorkbookPart workbookPart, string definedName)
        {
            string namedRange = workbookPart.Workbook.DefinedNames
                .FirstOrDefault(dn => ((DefinedName)dn).Name.Equals(definedName)).InnerText;

            if (namedRange == null)
                return "Defined name not found";

            string cellReference = namedRange;
            string[] parts = cellReference.Split('!');
            string sheetName = parts[0].Trim('\'');
            string cellAddress = parts[1];
            cellAddress = cellAddress.Replace("$", "");

            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>()
                .FirstOrDefault(s => s.Name == sheetName);

            if (sheet == null)
                return "Sheet not found";

            try
            {
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                Cell cell = worksheetPart.Worksheet.Descendants<Cell>()
                    .FirstOrDefault(c => c.CellReference == cellAddress);

                if (cell == null)
                {
                    throw new NullReferenceException();
                }

                return GetCellValue(workbookPart, cell);
            }
            catch (NullReferenceException e)
            {
                //MessageBox.Show(String.Format("Folgende Zelle {0} wurde nicht gefunden!", cellAddress), "Warnung");
                //MessageShow();
            }

            return "Cell not found";
        }

        private static string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            string value = cell.CellValue == null ? "" : cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return workbookPart.SharedStringTablePart.SharedStringTable.ChildElements.ElementAt(int.Parse(value)).InnerText;
            }
            return value;
        }
    }
}
