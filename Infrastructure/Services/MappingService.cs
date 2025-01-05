using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Core.Domain.Models;
using Core.Interfaces;

namespace Infrastructure.Services
{
    public class MappingService : IMappingService
    {
        private readonly IExcelService _excelService;
        private readonly IPowerPointService _pptService;

        public MappingService(IExcelService excelService, IPowerPointService pptService)
        {
            _excelService = excelService ?? throw new ArgumentNullException(nameof(excelService));
            _pptService = pptService ?? throw new ArgumentNullException(nameof(pptService));
        }

        public MappingResult ValidateMapping(string excelPath, string pptPath)
        {
            var excelNames = _excelService.GetDefinedNames();
            var pptNames = _pptService.GetShapeNames(pptPath);

            return FilterMissingAssignedNames(excelNames, pptNames);
        }

        public void ApplyMapping(string sourcePath, string targetPath)
        {
            var definedNames = _excelService.GetDefinedNames();

            foreach (var name in definedNames)
            {
                var value = _excelService.ExtractDataWithFieldName(name);
                _pptService.InsertDataIntoShape(name, value);
            }

            _pptService.SavePresentation(targetPath);
        }

        private MappingResult FilterMissingAssignedNames(IList<string> excelNames, IList<string> pptNames)
        {
            var powerPointMissingMapping = pptNames
                .Where(x => !excelNames.Contains(x))
                .ToList();

            var excelMissingMapping = excelNames
                .Where(x => !pptNames.Contains(x))
                .ToList();

            bool isMappingMatching = !powerPointMissingMapping.Any() &&
                                    !excelMissingMapping.Any();

            return new MappingResult(
                powerPointMissingMapping,
                excelMissingMapping,
                isMappingMatching);
        }
    }
}
