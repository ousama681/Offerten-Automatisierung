using System.Diagnostics;
using ShapeCrawler;
using UtilHelper;

namespace ExcelEditor
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string pptFile = "C:\\ZbwTechniker\\6tSemester\\Diplom Abschlussarbeit\\Entwicklung\\LatestOffertenAutomatisierung\\Offerten-Automatisierung\\IntegrationTests\\Testfiles\\Kundenofferte.pptx";
            string xlsFile = "C:\\ZbwTechniker\\6tSemester\\Diplom Abschlussarbeit\\Entwicklung\\LatestOffertenAutomatisierung\\Offerten-Automatisierung\\IntegrationTests\\Testfiles\\TestFileKeyNames.xlsx";

            UtilHelper.UtilHelper.ProcessPowerpoint(pptFile, xlsFile);

            //// Open the existing presentation
            //var presentation = new Presentation("C:\\ZbwTechniker\\6tSemester\\Diplom Abschlussarbeit\\20240904_Budget Kalt Cheese Technology.pptx");

            //foreach (var slide in presentation.Slides)
            //{

            //    foreach (var shape in slide.Shapes)
            //    {
            //        if (shape.Name.Equals("MachineType"))
            //        {
            //            shape.TextBox!.Text = "Dies ist ein Platzhaltertext um das bearbeiten per code zu testen.";
            //        }
            //    }

            //}
            //presentation.Save();
        }
    }
}
