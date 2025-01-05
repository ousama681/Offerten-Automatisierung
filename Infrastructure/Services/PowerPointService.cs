using Core.Interfaces;
using ShapeCrawler;

namespace Infrastructure.Services
{
    public class PowerPointService : IPowerPointService
    {
        private Presentation _presentation;

        public PowerPointService()
        {
        }
        public void LoadPowerPointFile(string path)
        {
            _presentation = new Presentation(path);
        }

        public IList<string> GetShapeNames(string path)
        {

            _presentation = new Presentation(path);

            return _presentation.Slides
                .SelectMany(slide => slide.Shapes)
                .Where(shape => !string.IsNullOrEmpty(shape.Name) && !shape.Name.Contains(" "))
                .Select(shape => shape.Name)
                .ToList();
        }

        public void InsertDataIntoShape(
string shapeName, string value)
        {
            foreach (var slide in _presentation.Slides)
            {
                foreach (var shape in slide.Shapes)
                {
                    if (shape.Name.Equals(shapeName))
                    {
                        shape.TextBox!.Text = value;
                        return;
                    }
                }
            }
        }

        public void SavePresentation(string targetPath)
        {
            _presentation.SaveAs(targetPath);
        }
    }
}
