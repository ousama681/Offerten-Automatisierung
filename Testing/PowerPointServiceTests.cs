using Core.Interfaces;
using Infrastructure.Services;

namespace Testing
{
    public class PowerPointServiceTests
    {
        private const string TEMPLATE_PATH = "TestData/template.pptx";
        private const string OUTPUT_PATH = "TestData/output.pptx";
        private IPowerPointService _pptService;

        [SetUp]
        public void Setup()
        {
            File.Copy(TEMPLATE_PATH, OUTPUT_PATH, true);
            _pptService = new PowerPointService();
        }

        [Test]
        public void InsertDataIntoShape_WithValidShape_UpdatesText()
        {
            const string shapeName = "TestShape";
            const string testValue = "Test Value";

       

            _pptService.LoadPowerPointFile(OUTPUT_PATH);
            _pptService.InsertDataIntoShape(shapeName, testValue);
            _pptService.SavePresentation(OUTPUT_PATH);

            Dictionary<string, string> shapeValuesAfter = _pptService.GetShapeValues(OUTPUT_PATH);

            string testShapeValueAfter = "";                
            shapeValuesAfter.TryGetValue(shapeName, out testShapeValueAfter);

            Assert.AreEqual(testShapeValueAfter, testValue);
        }

        [Test]
        public void GetShapeNames_ReturnsAllValidShapes()
        {
            var shapes = _pptService.GetShapeNames(TEMPLATE_PATH);
            Assert.That(shapes, Is.Not.Empty);
        }

        [TearDown]
        public void Cleanup()
        {
            (_pptService as IDisposable)?.Dispose();
            if (File.Exists(OUTPUT_PATH))
                File.Delete(OUTPUT_PATH);
        }
    }
}
