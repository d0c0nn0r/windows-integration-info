using CsvHelper.Configuration;
using WinDevOps.WMI;

namespace win_int_info.Serialization
{
    internal sealed class ProcPropMap : CsvClassMap<ProcessorProperties>
    {
        public ProcPropMap()
        {
            this.AutoMap();
            this.Map(m => m.Miscellaneous).Name("Miscellaneous").TypeConverter<EnumerableConverter>();

        }
    }
}