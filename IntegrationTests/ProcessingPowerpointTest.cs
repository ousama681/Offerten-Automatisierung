using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ShapeCrawler;
using UtilHelper;

namespace IntegrationTests
{
    [TestFixture]
    public class ProcessingPowerpointTest
    {


        [Test]
        public void CheckIfNewPowerpointFileWasCreated()
        {
            // Arrange
            string pptFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles", "Kundenofferte.pptx");
            string xlsFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles", "TestFileKeyNames.xlsx");
            string newPptFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles", "Kundenofferte_Processed.pptx");

            // Act
            string savedPptFile = UtilHelper.OfficeEditor.CreateNewPresentation(pptFile, xlsFile, newPptFile);


            //string newPptFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles", "Kundenofferte_Processed.pptx");
            // Assert
            // Check if new pptfile was created
            bool isCreated =  Path.Exists(savedPptFile);

            Assert.IsTrue(isCreated);
        }

        [Test]
        public void InsertMappedKeyValuesIntoPowerpoint()
        {
            // Arrange
            string pptFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles", "Kundenofferte.pptx");
            string xlsFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles", "TestFileKeyNames.xlsx");
            string newPptFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles", "Kundenofferte_Processed.pptx");

            // Act
            string savedPptFile = UtilHelper.OfficeEditor.CreateNewPresentation(pptFile, xlsFile, newPptFile);


            IList<string> keynames = GetListOfKeyNames(xlsFile);

            IDictionary<string, string> xlsKeyValues = GetCellValuesFromExcel(xlsFile, keynames);
            IDictionary<string, string> pptKeyValues = GetTextFieldValuesFromEditedPowerpoint(savedPptFile, xlsFile, keynames);
            /** Abfolge für Test
             * 1. Keyname saus Excel Values extrahieren
             * 2. Powerpoint bearbeiten
             * 3. Mit Keyname Wert aus einem beliebigne Textfeld aus neuer Powerpoint extrahieren.
             * 4. Wert mit ZellenWert aus Excel mitslebiugem Keyname vergleichen.
             * 
             * 
             */

            // Assert
            CollectionAssert.AreEquivalent(xlsKeyValues, pptKeyValues);
        }











        // HELPER METODS -----------------------------------------------------------------------
        public IDictionary<string, string> GetTextFieldValuesFromEditedPowerpoint(string pptFile, string xlsFile, IList<string> mappingKeys)
        {
            IDictionary<string, string> pptKeyValuePairs = new Dictionary<string, string>();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(xlsFile, false))
            {

                WorkbookPart workbookPart = doc.WorkbookPart!;
                DefinedNames definedNames = workbookPart.Workbook.DefinedNames!;



                var presentation = new Presentation(pptFile);


                foreach (DefinedName keyname in definedNames)
                {
                    foreach (var slide in presentation.Slides)
                    {

                        foreach (var shape in slide.Shapes)
                        {
                            if (shape.Name.Equals(keyname.Name))
                            {
                                string txtFieldValue = shape.TextBox!.Text;

                                pptKeyValuePairs.Add(keyname.Name!, txtFieldValue);
                            }
                        }

                    }
                }
            }

            return pptKeyValuePairs;
        }


        public Dictionary<string, string> GetCellValuesFromExcel(string fileName, IList<string> mappingKeys)
        {
            Dictionary<string, string> mappedValues = new Dictionary<string, string>();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))

                foreach (var key in mappingKeys)
                {
                    Sheet sheet = doc.WorkbookPart!.Workbook.Sheets!.GetFirstChild<Sheet>()!;
                    Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id!.Value!) as WorksheetPart)!.Worksheet;
                    IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>()!.Descendants<Row>();
                    IEnumerable<Cell> cells = worksheet.GetFirstChild<SheetData>()!.Descendants<Cell>();


                    WorkbookPart workbookPart = doc.WorkbookPart;
                    DefinedNames definedNames = workbookPart.Workbook.DefinedNames!;

                    if (definedNames == null)
                    {
                        throw new ArgumentNullException("There are no unique names defined for cells in the excel sheet.");
                    }

                    foreach (DefinedName dn in definedNames)
                    {
                        if (dn.Name!.Equals(key))
                        {
                            string cellReference = dn.Name!;
                            string cellValue = UtilHelper.OfficeEditor.ExtractDataWithFieldName(workbookPart, dn.Name!);
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

        public IList<string> GetListOfKeyNames(string filename)
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filename, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart!;
                DefinedNames definedNames = workbookPart.Workbook.DefinedNames!;
                IList<string> keynames = new List<string>();

                if (definedNames != null)
                {

                    foreach (DefinedName keyname in definedNames)
                    {
                        keynames.Add(keyname.Name!.ToString()!);
                    }
                }
                return keynames;
            }
        }
    }
}
