using ShapeCrawler;

namespace PowerPointEditor
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");


            // Find Powerpoint with Path as Param
            // open presentation and get first slide
            var pres = new Presentation("C:\\Projects\\PrototypDA\\PowerPointEditor\\helloWorld.pptx");
            var slide = pres.Slides.First();

            // get text holder auto shape
            var shape = slide.Shapes.First();

            // update text of shape
            //shape.TextBox.Text = "A new shape text";

            // change text for a certain paragraph
            var paragraph = shape.TextBox.Paragraphs[1];
            paragraph.Text = "A new text for second paragraph";

            // get font name and size
            var paragraphPortion = shape.TextBox.Paragraphs.First().Portions.First();
            Console.WriteLine($"Font name: {paragraphPortion.Font.LatinName}");
            Console.WriteLine($"Font size: {paragraphPortion.Font.Size}");

            // set bold font
            //paragraphPortion.Font.IsBold = true;

            // get font ARGB color
            var fontColor = paragraphPortion.Font.Color.Hex;

            // update font color
            paragraphPortion.Font.Color.Update("FF0000");
            paragraphPortion.Font.Size = 24;

            pres.Save();



            // Edit Textfield with a test String





            // Save Powerpoint
        }
    }
}
