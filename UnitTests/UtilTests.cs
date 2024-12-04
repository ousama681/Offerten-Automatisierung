using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using UtilHelper;

namespace UnitTests
{
    [TestFixture]
    public class UtilTests
    {
        [Test]
        public void TestCreateOutputForMappingfile()
        {
            // Arrange

            IList<string> keynames = new List<string>{
                "MachineType",
                "LitreAmount",
                "DurationDays",
                "TotalPrice",
                "BatchInformation"
            };

            // act
            string result = UtilHelper.UtilHelper.CreateOutputforMappingFile(keynames);

            // Assert
            Assert.AreEqual(result, "MachineType;LitreAmount;DurationDays;TotalPrice;BatchInformation;");
        }


        [Test]
        public void TestGetDefinedCellsValues()
        {
            // Arrange
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles", "TestFileKeyNames.xlsx");

            IList<string> keynames = new List<string>{
                "MachineType",
                "LitreAmount",
                "DurationDays",
                "TotalPrice",
                "BatchInformation"
            };

            // Act
            IList<string> result = UtilHelper.UtilHelper.GetListOfKeyNames(filePath);

            // Assert
            // The benfit of AreEquivalent() is the ability to check the existence of the strings while the order isnt important.
            CollectionAssert.AreEquivalent(keynames, result);
        }
    }
}
