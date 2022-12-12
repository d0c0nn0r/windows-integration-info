using CsvHelper.Configuration;
using WinDevOps.WMI;

namespace win_int_info.Serialization
{
    internal sealed class MemoryPropMap : CsvClassMap<MemoryProperties>
    {
        public MemoryPropMap()
        {
            this.AutoMap();
            this.Map(m => m.Miscellaneous).Name("Miscellaneous").TypeConverter<EnumerableConverter>();

        }
    }
}