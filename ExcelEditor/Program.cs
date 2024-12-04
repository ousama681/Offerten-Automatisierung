using System.Diagnostics;
using ShapeCrawler;

namespace ExcelEditor
{
    internal class Program
    {
        static void Main(string[] args)
        {

            // Open the existing presentation
            var presentation = new Presentation("C:\\ZbwTechniker\\6tSemester\\Diplom Abschlussarbeit\\20240904_Budget Kalt Cheese Technology.pptx");

            foreach (var slide in presentation.Slides)
            {

                foreach (var shape in slide.Shapes)
                {
                    if (shape.Name.Equals("MachineType"))
                    {
                        shape.TextBox!.Text = "Dies ist ein Platzhaltertext um das bearbeiten per code zu testen.";
                    }
                }
      
            }
            presentation.Save();
        }
    }
}
