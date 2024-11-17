using static ExcelEditor.Helper;

namespace UnitTests
{
    [TestFixture]
    public class ExcelTasksTest
    {
        [Test]
        public void TestGetDefinedCellsValues()
        {
            // Arrange
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Testfiles", "TestFile.xlsx");

            List<Mapping> definedCells = new List<Mapping>{
                new Mapping("TextFieldA", "A1"),
                new Mapping("TextFieldB", "D27"),
                new Mapping("TextFieldC", "I16"),
                new Mapping("TextFieldD", "M23")
            };

            // Act
            string result = ExcelEditor.Helper.GetDefinedCellsValues(filePath, definedCells);

            // Assert
            Assert.AreEqual(result, "A1: Press ==> TextFieldA\r\nD27: Parallel ==> TextFieldB\r\nI16: Round ==> TextFieldC\r\nM23: 50000 ==> TextFieldD\r\n");
        }
    }
}