using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
