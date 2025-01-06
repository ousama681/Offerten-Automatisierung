using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Core.Domain.Models;

namespace Core.Interfaces
{
    public interface IMappingService
    {
        MappingResult ValidateMapping(string excelPath, string pptPath);
        void ApplyMapping(string sourcePath, string targetPath);
    }
}
