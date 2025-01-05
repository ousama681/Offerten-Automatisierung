using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Core.Domain.Models
{
    public class MappingResult
    {
        public IReadOnlyList<string> ExcelMissingMapping { get; }
        public IReadOnlyList<string> PowerPointMissingMapping { get; }
        public bool IsMappingMatching { get; }

        public MappingResult(
            IEnumerable<string> powerPointMissingMapping,
            IEnumerable<string> excelMissingMapping,
            bool isMappingMatching)
        {
            PowerPointMissingMapping = powerPointMissingMapping?.ToList() ?? new List<string>();
            ExcelMissingMapping = excelMissingMapping?.ToList() ?? new List<string>();
            IsMappingMatching = isMappingMatching;
        }

        public bool HasMissingMappings =>
            PowerPointMissingMapping.Any() || ExcelMissingMapping.Any();
    }
}
