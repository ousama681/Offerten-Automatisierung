using Core.Interfaces;
using Infrastructure.Services;

namespace Testing
{
    public class ExcelServiceTests
    {
        private const string TEST_EXCEL_PATH = "TestData/test_extraction.xlsx";
        private IExcelService _excelService;

        [SetUp]
        public void Setup()
        {
            _excelService = new ExcelService();
            _excelService.LoadExcelFile(TEST_EXCEL_PATH);
        }

        [Test]
        [TestCase("NamedCell1", "ExpectedValue1")]
        [TestCase("EmptyCell", "Cell 'E1' has no Value in it!")]
        public void ExtractDataWithFieldName_WithValidNames_ReturnsExpectedValues(string fieldName, string expected)
        {
            var result = _excelService.ExtractDataWithFieldName(fieldName);
            Assert.That(result, Is.EqualTo(expected));
        }

        [Test]
        public void ExtractDataWithFieldName_WithInvalidName_ThrowsException()
        {
            Assert.Throws<InvalidOperationException>(() =>
                _excelService.ExtractDataWithFieldName("NonExistentField"));
        }

        [Test]
        public void GetDefinedNames_ReturnsAllValidNames()
        {
            var names = _excelService.GetDefinedNames();
            Assert.That(names, Is.Not.Empty);
            Assert.That(names, Does.Not.Contain(null));
            Assert.That(names, Does.Not.Contain(string.Empty));
        }

        [TearDown]
        public void Cleanup()
        {
            (_excelService as IDisposable)?.Dispose();
        }
    }
}
