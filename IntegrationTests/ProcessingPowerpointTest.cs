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

            // Act
            string newPptFile = UtilHelper.UtilHelper.ProcessPowerpoint(pptFile, xlsFile);


            //string newPptFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles", "Kundenofferte_Processed.pptx");
            // Assert
            // Check if new pptfile was created
            bool isCreated =  Path.Exists(newPptFile);

            Assert.IsTrue(isCreated);
        }

        [Test]
        public void InsertMappedKeyValuesIntoPowerpoint()
        {
            // Arrange
            string pptFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles", "Kundenofferte.pptx");
            string xlsFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles", "TestFileKeyNames.xlsx");

            // Act
            string newPptFile = UtilHelper.UtilHelper.ProcessPowerpoint(pptFile, xlsFile);


            IList<string> keynames = UtilHelper.UtilHelper.GetListOfKeyNames(xlsFile);

            IDictionary<string, string> xlsKeyValues = UtilHelper.UtilHelper.GetCellValuesFromExcel(xlsFile, keynames);
            IDictionary<string, string> pptKeyValues = GetTextFieldValuesFromEditedPowerpoint(newPptFile, xlsFile, keynames);
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
        public static IDictionary<string, string> GetTextFieldValuesFromEditedPowerpoint(string pptFile, string xlsFile, IList<string> mappingKeys)
        {
            IDictionary<string, string> pptKeyValuePairs = new Dictionary<string, string>();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(xlsFile, false))
            {

                WorkbookPart workbookPart = doc.WorkbookPart;
                DefinedNames definedNames = workbookPart.Workbook.DefinedNames;



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

                                pptKeyValuePairs.Add(keyname.Name, txtFieldValue);
                            }
                        }

                    }
                }
            }

            return pptKeyValuePairs;
        }




    }
}
