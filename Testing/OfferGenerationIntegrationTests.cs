using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Core.Interfaces;
using Infrastructure.Services;

namespace Testing
{
    public class OfferGenerationIntegrationTests
    {
        private const string TEST_EXCEL_PATH = "TestData/test_data_types.xlsx";
        private const string TEMPLATE_PATH = "TestData/template.pptx";
        private const string OUTPUT_PATH = "TestData/generated_offer.pptx";

        private IExcelService _excelService;
        private IPowerPointService _pptService;
        private IMappingService _mappingService;

        [SetUp]
        public void Setup()
        {
            _excelService = new ExcelService();
            _pptService = new PowerPointService();
            _mappingService = new MappingService(_excelService, _pptService);
        }

        [Test]
        public void CompleteOfferGeneration_WithValidData_GeneratesCorrectOffer()
        {
            // TC009: Kompletter Offertengenerierungsprozess
            _mappingService.ApplyMapping(TEST_EXCEL_PATH, OUTPUT_PATH);
            Assert.That(File.Exists(OUTPUT_PATH));
        }

        [Test]
        public void DataTypeHandling_WithVariousTypes_MapsCorrectly()
        {
            // TC003: Verschiedene Datentypen
            var definedNames = _excelService.GetDefinedNames();
            foreach (var name in definedNames)
            {
                var value = _excelService.ExtractDataWithFieldName(name);
                Assert.That(value, Is.Not.Null);
            }
        }

        [Test]
        public void PerformanceTest_WithLargeDataset_CompletesWithinTimeLimit()
        {
            // TC010: Verarbeitungszeit
            var stopwatch = Stopwatch.StartNew();

            _mappingService.ApplyMapping(TEST_EXCEL_PATH, OUTPUT_PATH);

            stopwatch.Stop();
            Assert.That(stopwatch.ElapsedMilliseconds, Is.LessThan(30000));
        }

        [Test]
        public void ResourceUsage_WithMultipleExecutions_MaintainsStableMemory()
        {
            // TC011: Ressourcenverbrauch
            long initialMemory = GC.GetTotalMemory(true);

            for (int i = 0; i < 5; i++)
            {
                _mappingService.ApplyMapping(TEST_EXCEL_PATH, $"output_{i}.pptx");
            }

            GC.Collect();
            long finalMemory = GC.GetTotalMemory(true);

            Assert.That(finalMemory - initialMemory, Is.LessThan(10_000_000)); // 10MB Toleranz
        }

        [Test]
        public void FormatPreservation_WithFormattedTemplate_PreservesFormatting()
        {
            // TC005: Formaterhaltung
            const string formattedTemplatePath = "TestData/formatted_template.pptx";
            var originalService = new PowerPointService();
            var originalShapes = originalService.GetShapeNames();

            _mappingService.ApplyMapping(TEST_EXCEL_PATH, OUTPUT_PATH);

            var resultService = new PowerPointService();
            var resultShapes = resultService.GetShapeNames();

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
