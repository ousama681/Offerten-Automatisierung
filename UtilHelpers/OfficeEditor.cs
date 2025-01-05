using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ShapeCrawler;

namespace UtilHelper
{
    public class OfficeEditor
    {
        public static string ExtractDataWithFieldName(WorkbookPart workbookPart, string keyname)
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

            return ExtractData(workbookPart, cell);
        }

        private static string ExtractData(WorkbookPart workbookPart, Cell cell)
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
                cellFormat = (CellFormat)workbookPart.WorkbookStylesPart!.Stylesheet.CellFormats!
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


        public static void InsertDataIntoShape(Presentation presentation, DefinedNames definedNames, WorkbookPart workbookPart)
        {
            foreach (DefinedName keyname in definedNames)
            {
                foreach (var slide in presentation.Slides)
                {
                    foreach (var shape in slide.Shapes)
                    {
                        if (shape.Name.Equals(keyname.Name))
                        {
                            shape.TextBox!.Text = ExtractDataWithFieldName(workbookPart, keyname.Name!);
                        }
                    }
                }
            }
        }



        public static string CreateNewPresentation(string pptFile, string xlsFile, string savePath)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(xlsFile, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart!;
                DefinedNames definedNames = workbookPart.Workbook.DefinedNames!;
                IList<string> keynames = new List<string>();

                if (definedNames != null)
                {
                    var presentation = new Presentation(pptFile);

                    InsertDataIntoShape(presentation, definedNames, workbookPart);
                    presentation.SaveAs(savePath);

                    return savePath;
                }
            }
            return "";
        }

        //public static bool isDefinedNamesEqual(IList<string> definedNamesExcel, IList<string> definedNamesPowerPoint)
        public static MappingResult IsFieldNamesEqual(string pptFile, string xlsFile)
        {
            IList<string> definedNamesExcel = new List<string>();
            IList<string> definedNamesPowerPoint = GetPowerPointFieldNames(pptFile);

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(xlsFile, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart!;
                DefinedNames definedNames = workbookPart.Workbook.DefinedNames!;

                if (definedNames != null)
                {
                    definedNamesExcel = definedNames.Select(dn => ((DefinedName)dn).Name!.ToString()).ToList()!;
                }
            }
            return FilterMissingAssignedNames(definedNamesExcel, definedNamesPowerPoint);
        }

        private static MappingResult FilterMissingAssignedNames(IList<string> excelNames, IList<string> pptNames)
        {
            IList<string> powerPointMissingMapping = pptNames.Where(x => !excelNames.Contains(x)).ToList();
            IList<string> excelMissingMapping = excelNames.Where(x => !pptNames.Contains(x)).ToList();

            bool isMappingMatching = !powerPointMissingMapping.Any() && !excelMissingMapping.Any();

            var result = new MappingResult(powerPointMissingMapping, excelMissingMapping, isMappingMatching);

            return result;
        }

        public static List<string> GetPowerPointFieldNames(string pptfile)
        {
            // Open the presentation
            var presentation = new Presentation(pptfile);

            // Get all shapes from all slides
            var allShapes = presentation.Slides.SelectMany(slide => slide.Shapes);

            // Filter shapes with non-empty Names and create a list of those names
            List<string> shapeNames = allShapes
                .Where(shape => !string.IsNullOrEmpty(shape.Name) && !shape.Name.Contains(" "))
                .Select(shape => shape.Name)
                .ToList();

            return shapeNames;
        }


        public class MappingResult
        {
            public IList<string> ExcelMissingMapping { get; set; }
            public IList<string> PowerPointMissingMapping { get; set; }

            public bool IsMappingMatching { get; set; }
            public MappingResult(IList<string> powerPointMissingMapping, IList<string> excelMissingMapping, bool isMappingMatching)
            {
                PowerPointMissingMapping = powerPointMissingMapping;
                ExcelMissingMapping = excelMissingMapping;
                IsMappingMatching = isMappingMatching;
            }
        }
    }
}
