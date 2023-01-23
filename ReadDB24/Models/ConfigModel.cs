using System.Collections.Generic;

namespace ReadDB24.Models
{
    internal class ConfigModel
    {
        public List<MappingExcelModel> Main { get; set; }
        public List<MappingExcelModel> Attached { get; set; }
        public List<MappingExcelModel> Reserved { get; set; }
    }
}
