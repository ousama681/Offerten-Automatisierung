using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ShapeCrawler;

namespace UtilHelper
{
    public class UtilHelper
    {


        public static string CreateOutputforMappingFile(IList<string> keynames)
        {
            string mapping = "";

            foreach (var keyname in keynames)
            {
                mapping += keyname.ToString() + ";";
            }

            return mapping;
        }

        public static IList<string> GetListOfKeyNames(string filename)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filename, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                DefinedNames definedNames = workbookPart.Workbook.DefinedNames;
                IList<string> keynames = new List<string>();

                if (definedNames != null)
                {

                    foreach (DefinedName keyname in definedNames)
                    {
                        keynames.Add(keyname.Name.ToString());
                    }
                }
                return keynames;
            }
        }

        public static IDictionary<string, string> CreateKeyValueMapping(string filename, IList<string> keynames)
        {
            Dictionary<string, string> keyvalueMapping = new Dictionary<string, string>();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filename, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;

                if (workbookPart != null)
                {
                    foreach (var keyname in keynames)
                    {
                        string output = GetCellValueFromDefinedName(workbookPart, keyname);

                        keyvalueMapping.Add(keyname, output);
                    }
                }
            }
            return keyvalueMapping;
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

        public static string GetCellValueFromDefinedName(WorkbookPart workbookPart, string keyname)
        {
            string namedRange = workbookPart.Workbook.DefinedNames
                .FirstOrDefault(dn => ((DefinedName)dn).Name.Equals(keyname)).InnerText;

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


        public static void InsertInputFromExcel(Presentation presentation, string keyword, string input)
        {
            foreach (var slide in presentation.Slides)
            {
                foreach (var shape in slide.Shapes)
                {
                    if (shape.Name.Equals(keyword))
                    {
                        shape.TextBox!.Text = input;
                    }
                }
            }
        }


        //public static void editSecondParagraphOfTextBox()
        //{


        //    Console.WriteLine("Hello, World!");


        //    // Find Powerpoint with Path as Param
        //    // open presentation and get first slide
        //    var pres = new Presentation("C:\\Projects\\PrototypDA\\PowerPointEditor\\helloWorld.pptx");
        //    var slide = pres.Slides.First();

        //    // get text holder auto shape
        //    var shape = slide.Shapes.First();

        //    // A Paragraph is basically the second Line of a Text in a Textbox, only if u pressed Enter.

        //    // change text for a certain paragraph
        //    var paragraph = shape.TextBox.Paragraphs[1];
        //    paragraph.Text = "A new text for second paragraph";

        //    // poritons.first greift auf den ersten Character eines Paragraphs zu.
        //    // get font name and size
        //    var paragraphPortion = shape.TextBox.Paragraphs.First().Portions.ElementAt(2);
        //    Console.WriteLine($"Font name: {paragraphPortion.Font.LatinName}");
        //    Console.WriteLine($"Font size: {paragraphPortion.Font.Size}");

        //    // set bold font
        //    paragraphPortion.Font.IsBold = true;

        //    // get font ARGB color
        //    var fontColor = paragraphPortion.Font.Color.Hex;

        //    // update font color
        //    paragraphPortion.Font.Color.Update("FF0000");

        //    pres.Save();
        //}


        //Console.WriteLine("Hello, World!");


        //    // Find Powerpoint with Path as Param
        //    // open presentation and get first slide
        //    var pres = new Presentation("C:\\Projects\\PrototypDA\\PowerPointEditor\\helloWorld.pptx");
        //var slide = pres.Slides.First();

        //// get text holder auto shape
        //var shape = slide.Shapes.First();

        //// update text of shape
        ////shape.TextBox.Text = "A new shape text";

        //// change text for a certain paragraph
        //var paragraph = shape.TextBox.Paragraphs[1];
        //paragraph.Text = "A new text for second paragraph";

        //    // get font name and size
        //    var paragraphPortion = shape.TextBox.Paragraphs.First().Portions.First();
        //Console.WriteLine($"Font name: {paragraphPortion.Font.LatinName}");
        //    Console.WriteLine($"Font size: {paragraphPortion.Font.Size}");

        //    // set bold font
        //    //paragraphPortion.Font.IsBold = true;

        //    // get font ARGB color
        //    var fontColor = paragraphPortion.Font.Color.Hex;

        //// update font color
        //paragraphPortion.Font.Color.Update("FF0000");
        //    paragraphPortion.Font.Size = 24;

        //    pres.Save();



        //    // Edit Textfield with a test String





        //    // Save Powerpoint

    }
}
