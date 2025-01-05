using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Core.Interfaces
{
    public interface IExcelService
    {
        void LoadExcelFile(string path);
        string ExtractDataWithFieldName(string keyname);
        IList<string> GetDefinedNames();
        Dictionary<string, string> GetDefinedNamesWithValues(string path);
    }
}
