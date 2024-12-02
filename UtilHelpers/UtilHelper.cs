using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ShapeCrawler;

namespace UtilHelper
{
    public class UtilHelper
    {
        public static string CreateOutputforMappingFile(List<string> items)
        {
            string mappings = "";

            foreach (var mapping in items)
            {
                mappings += mapping.ToString() + ";";
            }

            return mappings;
        }


        public static void editTextBoxText()
        {
            var pres = new Presentation("C:\\Projects\\PrototypDA\\PowerPointEditor\\helloWorld.pptx");
            var slide = pres.Slides.First();

            // get text holder auto shape
            var shape = slide.Shapes.First();

            // update text of shape
            shape.TextBox.Text = "A new shape text";
            pres.Save();
        }


        public static void editSecondParagraphOfTextBox()
        {


            Console.WriteLine("Hello, World!");


            // Find Powerpoint with Path as Param
            // open presentation and get first slide
            var pres = new Presentation("C:\\Projects\\PrototypDA\\PowerPointEditor\\helloWorld.pptx");
            var slide = pres.Slides.First();

            // get text holder auto shape
            var shape = slide.Shapes.First();

            // A Paragraph is basically the second Line of a Text in a Textbox, only if u pressed Enter.

            // change text for a certain paragraph
            var paragraph = shape.TextBox.Paragraphs[1];
            paragraph.Text = "A new text for second paragraph";

            // poritons.first greift auf den ersten Character eines Paragraphs zu.
            // get font name and size
            var paragraphPortion = shape.TextBox.Paragraphs.First().Portions.ElementAt(2);
            Console.WriteLine($"Font name: {paragraphPortion.Font.LatinName}");
            Console.WriteLine($"Font size: {paragraphPortion.Font.Size}");

            // set bold font
            paragraphPortion.Font.IsBold = true;

            // get font ARGB color
            var fontColor = paragraphPortion.Font.Color.Hex;

            // update font color
            paragraphPortion.Font.Color.Update("FF0000");

            pres.Save();
        }


        //Console.WriteLine("Hello, World!");


        //    // Find Powerpoint with Path as Param
        //    // open presentation and get first slide
        //    var pres = new Presentation("C:\\Projects\\PrototypDA\\PowerPointEditor\\helloWorld.pptx");
        //var slide = pres.Slides.First();

        //// get text holder auto shape
        //var shape = slide.Shapes.First();

        //// update text of shape
        ////shape.TextBox.Text = "A new shape text";

        //// change text for a certain paragraph
        //var paragraph = shape.TextBox.Paragraphs[1];
        //paragraph.Text = "A new text for second paragraph";

        //    // get font name and size
        //    var paragraphPortion = shape.TextBox.Paragraphs.First().Portions.First();
        //Console.WriteLine($"Font name: {paragraphPortion.Font.LatinName}");
        //    Console.WriteLine($"Font size: {paragraphPortion.Font.Size}");

        //    // set bold font
        //    //paragraphPortion.Font.IsBold = true;

        //    // get font ARGB color
        //    var fontColor = paragraphPortion.Font.Color.Hex;

        //// update font color
        //paragraphPortion.Font.Color.Update("FF0000");
        //    paragraphPortion.Font.Size = 24;

        //    pres.Save();



        //    // Edit Textfield with a test String





        //    // Save Powerpoint

    }
}
