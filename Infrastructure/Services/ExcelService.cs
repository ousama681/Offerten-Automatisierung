using Core.Interfaces;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Infrastructure.Services
{
    public class ExcelService : IExcelService
    {
        private WorkbookPart _workbookPart;
        private SpreadsheetDocument _spreadsheetDocument;


        public ExcelService()
        {
        }

        public void LoadExcelFile(string path)
        {
            if (_spreadsheetDocument != null)
            {
                _spreadsheetDocument.Dispose();
            }

            _spreadsheetDocument = SpreadsheetDocument.Open(path, false);
            _workbookPart = _spreadsheetDocument.WorkbookPart
                ?? throw new InvalidOperationException("Excel file has no workbook part.");
        }

        public string ExtractDataWithFieldName(string keyname)
        {
            ArgumentNullException.ThrowIfNull(keyname);

            var namedRange = _workbookPart.Workbook?.DefinedNames?
                .FirstOrDefault(dn => ((DefinedName)dn).Name!.Equals(keyname))?.InnerText
                ?? throw new InvalidOperationException($"Defined name '{keyname}' not found");

            string[] parts = namedRange.Split('!');
            if (parts.Length != 2)
                throw new FormatException("Invalid named range format");

            string sheetName = parts[0].Trim('\'');
            string cellAddress = parts[1].Replace("$", "");

            var sheet = _workbookPart.Workbook.Descendants<Sheet>()
    .FirstOrDefault(s => s.Name?.Value == sheetName)
    ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found");

            var worksheetPart = (WorksheetPart)_workbookPart.GetPartById(sheet.Id!)
                ?? throw new InvalidOperationException("Worksheet part not found");

            var cell = worksheetPart.Worksheet.Descendants<Cell>()
                .FirstOrDefault(c => c.CellReference == cellAddress)
                ?? throw new InvalidOperationException($"Cell '{cellAddress}' not found");

            return ExtractCellValue(cell);
        }

        public IList<string> GetDefinedNames()
        {
            var definedNames = _workbookPart.Workbook.DefinedNames;
            if (definedNames == null)
                return new List<string>();

            return definedNames
                .Cast<DefinedName>()
                .Select(dn => dn.Name?.Value)
                .Where(name => !string.IsNullOrEmpty(name))
                .ToList();
        }


        private string ExtractCellValue(Cell cell)
        {
            if (cell.CellValue == null)
                return string.Empty;

            string value = cell.CellValue.InnerText;

            if (cell.DataType?.Value == CellValues.SharedString)
            {
                return _workbookPart.SharedStringTablePart!
                    .SharedStringTable
                    .ChildElements
                    .ElementAt(int.Parse(value))
                    .InnerText;
            }

            return FormatNumericValue(cell);
        }

        private string FormatNumericValue(Cell cell)
        {
            if (cell.StyleIndex == null || cell.CellValue == null)
                return cell.CellValue?.InnerText ?? string.Empty;

            var cellFormat = (CellFormat)_workbookPart.WorkbookStylesPart!
                .Stylesheet
                .CellFormats!
                .ElementAt((int)cell.StyleIndex.Value);

            if (cellFormat?.NumberFormatId == null)
                return cell.CellValue.InnerText;

            var numberingFormat = _workbookPart.WorkbookStylesPart!
                .Stylesheet
                .NumberingFormats?
                .Cast<NumberingFormat>()
                .FirstOrDefault(f => f.NumberFormatId?.Value == cellFormat.NumberFormatId.Value);

            if (numberingFormat?.FormatCode?.Value != null &&
                double.TryParse(cell.CellValue.InnerText, out double number))
            {
                return number.ToString(numberingFormat.FormatCode.Value);
            }

            return cell.CellValue.InnerText;
        }

        public void Dispose()
        {
            _workbookPart?.OpenXmlPackage?.Dispose();
        }
    }
}
