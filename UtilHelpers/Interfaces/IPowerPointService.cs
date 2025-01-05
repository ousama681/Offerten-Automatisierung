using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Core.Interfaces
{
    public interface IPowerPointService
    {
        void LoadPowerPointFile(string path);
        void InsertDataIntoShape(string shapeName, string value);
        IList<string> GetShapeNames(string path);
        void SavePresentation(string targetPath);
        Dictionary<string, string> GetShapeValues(string path);

    }
}
