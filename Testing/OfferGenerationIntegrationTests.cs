using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Core.Interfaces;
using DocumentFormat.OpenXml;
using Infrastructure.Services;

namespace Testing
{
    public class OfferGenerationIntegrationTests
    {
        private const string TEST_EXCEL_PATH = "TestData/test_data_types.xlsx";
        private const string TEMPLATE_PATH = "TestData/template_data_types.pptx";
        private const string OUTPUT_PATH = "TestData/generated_offer.pptx";

        private IExcelService _excelService;
        private IPowerPointService _pptService;
        private IMappingService _mappingService;

        [SetUp]
        public void Setup()
        {
            _excelService = new ExcelService();
            _excelService.LoadExcelFile(TEST_EXCEL_PATH);
            _pptService = new PowerPointService();
            _pptService.LoadPowerPointFile(TEMPLATE_PATH);
            _mappingService = new MappingService(_excelService, _pptService);
        }

        [Test]
        public void CompleteOfferGeneration_WithValidData_GeneratesCorrectOffer()
        {     
            _mappingService.ApplyMapping(TEST_EXCEL_PATH, OUTPUT_PATH);
            Assert.That(File.Exists(OUTPUT_PATH));
        }

        [Test]
        public void DataTypeHandling_WithVariousTypes_MapsCorrectly()
        {
            _mappingService.ApplyMapping(TEST_EXCEL_PATH, OUTPUT_PATH);


            Dictionary<string, string> insertedShapeValues = _pptService.GetShapeValues(OUTPUT_PATH);
            Dictionary<string, string> extractedFieldValues = _excelService.GetDefinedNamesWithValues(TEST_EXCEL_PATH);

            CollectionAssert.AreEquivalent(insertedShapeValues, extractedFieldValues);
        }

        [Test]
        public void FormatPreservation_WithFormattedTemplate_PreservesFormatting()
        {
            const string formattedTemplatePath = "TestData/formatted_template.pptx";
            var originalService = new PowerPointService();
            var originalShapes = originalService.GetShapeNames(TEMPLATE_PATH);

            _mappingService.ApplyMapping(TEST_EXCEL_PATH, OUTPUT_PATH);

            var resultService = new PowerPointService();
            var resultShapes = resultService.GetShapeNames(OUTPUT_PATH);

            Assert.That(resultShapes, Is.EquivalentTo(originalShapes));
        }

        [TearDown]
        public void Cleanup()
        {
            (_excelService as IDisposable)?.Dispose();
            (_pptService as IDisposable)?.Dispose();

            var outputFiles = Directory.GetFiles("TestData", "output_*.pptx");
            foreach (var file in outputFiles)
            {
                File.Delete(file);
            }
        }
    }
}
