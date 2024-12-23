using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ShapeCrawler;

namespace UtilHelper
{
    public class UtilHelper
    {


        //public static string CreateOutputforMappingFile(IList<string> keynames)
        //{
        //    string mapping = "";

        //    foreach (var keyname in keynames)
        //    {
        //        mapping += keyname + ";";
        //    }

        //    return mapping;
        //}

        //public static IDictionary<string, string> CreateKeyValueMapping(string filename, IList<string> keynames)
        //{
        //    Dictionary<string, string> keyvalueMapping = new Dictionary<string, string>();

        //    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filename, false))
        //    {
        //        WorkbookPart workbookPart = doc.WorkbookPart;

        //        if (workbookPart != null)
        //        {
        //            foreach (var keyname in keynames)
        //            {
        //                string output = GetCellValueFromDefinedName(workbookPart, keyname);

        //                keyvalueMapping.Add(keyname, output);
        //            }
        //        }
        //    }
        //    return keyvalueMapping;
        //}

        //public static Dictionary<string, string> GetCellValuesFromExcel(string fileName, IList<string> mappingKeys)
        //{
        //    Dictionary<string, string> mappedValues = new Dictionary<string, string>();

        //    using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))

        //        foreach (var key in mappingKeys)
        //        {
        //            Sheet sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
        //            Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
        //            IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
        //            IEnumerable<Cell> cells = worksheet.GetFirstChild<SheetData>().Descendants<Cell>();


        //            WorkbookPart workbookPart = doc.WorkbookPart;
        //            DefinedNames definedNames = workbookPart.Workbook.DefinedNames;

        //            if (definedNames == null)
        //            {
        //                throw new ArgumentNullException("There are no unique names defined for cells in the excel sheet.");
        //            }

        //            foreach (DefinedName dn in definedNames)
        //            {
        //                if (dn.Name.Equals(key))
        //                {
        //                    string cellReference = dn.Name;
        //                    string cellValue = GetCellValueFromDefinedName(workbookPart, dn.Name);
        //                    mappedValues.Add(cellReference, cellValue);
        //                }
        //            }
        //        }


        //    if (mappedValues.Count == 0)
        //    {
        //        throw new ArgumentNullException("There were no matching defined names found in the excel sheet that matched the provided ones from the tool.");
        //    }

        //    return mappedValues;
        //}

        //public static string GetCellValueFromDefinedName(WorkbookPart workbookPart, string keyname)
        //{
        //    string namedRange = workbookPart.Workbook.DefinedNames
        //        .FirstOrDefault(dn => ((DefinedName)dn).Name.Equals(keyname)).InnerText;

        //    if (namedRange == null)
        //        return "Defined name not found";

        //    string cellReference = namedRange;
        //    string[] parts = cellReference.Split('!');
        //    string sheetName = parts[0].Trim('\'');
        //    string cellAddress = parts[1];
        //    cellAddress = cellAddress.Replace("$", "");

        //    Sheet sheet = workbookPart.Workbook.Descendants<Sheet>()
        //        .FirstOrDefault(s => s.Name == sheetName);

        //    if (sheet == null)
        //        return "Sheet not found";


        //    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        //    Cell cell = worksheetPart.Worksheet.Descendants<Cell>()
        //        .FirstOrDefault(c => c.CellReference == cellAddress);

        //    if (cell == null)
        //    {
        //        throw new NullReferenceException();
        //    }

        //    return GetCellValue(workbookPart, cell);
        //}

        public static string GetCellValueFromDefinedName(WorkbookPart workbookPart, string keyname)
        {
            ArgumentNullException.ThrowIfNull(workbookPart);
            ArgumentNullException.ThrowIfNull(keyname);

            var namedRange = workbookPart.Workbook?.DefinedNames?
                .FirstOrDefault(dn => ((DefinedName)dn).Name!.Equals(keyname))?.InnerText
                ?? throw new InvalidOperationException("Defined name not found");

            string[] parts = namedRange.Split('!');
            if (parts.Length != 2)
                throw new FormatException("Invalid named range format");

            string sheetName = parts[0].Trim('\'');
            string cellAddress = parts[1].Replace("$", "");

            var sheet = workbookPart.Workbook.Descendants<Sheet>()
                .FirstOrDefault(s => s.Name == sheetName)
                ?? throw new InvalidOperationException("Sheet not found");

            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!)
                ?? throw new InvalidOperationException("Worksheet part not found");

            var cell = worksheetPart.Worksheet.Descendants<Cell>()
                .FirstOrDefault(c => c.CellReference == cellAddress)
                ?? throw new InvalidOperationException("Cell not found");

            return GetCellValue(workbookPart, cell);
        }

        private static string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            string value = cell.CellValue == null ? "" : cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return workbookPart.SharedStringTablePart!.SharedStringTable.ChildElements.ElementAt(int.Parse(value)).InnerText;
            }
            return FormattedNumber(workbookPart, cell);
        }

        private static string FormattedNumber(WorkbookPart workbookPart, Cell cell)
        {
            if (cell.CellValue == null)
                return "";

            CellFormat cellFormat = null;
            if (cell.StyleIndex != null)
            {
                cellFormat = (CellFormat) workbookPart.WorkbookStylesPart!.Stylesheet.CellFormats!
                    .ElementAt((int)cell.StyleIndex.Value);
            }

            if (cellFormat == null || cellFormat.NumberFormatId == null)
                return cell.CellValue.InnerText;

            var numberingFormats = workbookPart.WorkbookStylesPart!.Stylesheet.NumberingFormats;
            if (numberingFormats != null)
            {
                var formatCode = numberingFormats.Elements<NumberingFormat>()
                    .Where(f => f.NumberFormatId!.Value == cellFormat.NumberFormatId.Value)
                    .Select(f => f.FormatCode!.Value)
                    .FirstOrDefault();

                if (formatCode != null)
                {
                    double number;
                    if (double.TryParse(cell.CellValue.InnerText, out number))
                    {
                        return number.ToString(formatCode);
                    }
                }
            }

            return cell.CellValue.InnerText;
        }


        public static void InsertInputFromExcel(Presentation presentation, DefinedNames definedNames, WorkbookPart workbookPart)
        {
            foreach (DefinedName keyname in definedNames)
            {
                foreach (var slide in presentation.Slides)
                {
                    foreach (var shape in slide.Shapes)
                    {
                        if (shape.Name.Equals(keyname.Name))
                        {
                            shape.TextBox!.Text = GetCellValueFromDefinedName(workbookPart, keyname.Name!);
                        }
                    }
                }
            }
        }



        public static string ProcessPowerpoint(string pptFile, string xlsFile, string savePath)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(xlsFile, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart!;
                DefinedNames definedNames = workbookPart.Workbook.DefinedNames!;
                IList<string> keynames = new List<string>();

                if (definedNames != null)
                {
                    var presentation = new Presentation(pptFile);

                    InsertInputFromExcel(presentation, definedNames, workbookPart);
                    presentation.SaveAs(savePath);

                    return savePath;
                }
            }
            return "";
        }
    }
}
