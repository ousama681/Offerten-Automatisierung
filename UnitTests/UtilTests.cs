using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using UtilHelper;

namespace UnitTests
{
    [TestFixture]
    public class UtilTests
    {
        [Test]
        public void TestCreateOutputForMappingfile()
        {
            // Arrange

            List<string> items = new List<string>();

            items.Add("A3 ==> TextFieldA");
            items.Add("D23 ==> TextFieldB");
            items.Add("G35 ==> TextFieldC");

            // act
            string result = UtilHelper.UtilHelper.CreateOutputforMappingFile(items);

            // Assert
            Assert.AreEqual(result, "A3 ==> TextFieldA;D23 ==> TextFieldB;G35 ==> TextFieldC;");
        }
    }
}
