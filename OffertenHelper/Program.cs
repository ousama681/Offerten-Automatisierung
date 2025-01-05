using Infrastructure.Services;

namespace Offerten_Helper
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            var excelService = new ExcelService();
            var pptService = new PowerPointService();
            var mappingService = new MappingService(excelService, pptService);
            Application.Run(new AutomatedOffer(excelService, pptService, mappingService));
        }
    }
}