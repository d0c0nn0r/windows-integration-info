using CsvHelper.Configuration;
using WinDevOps.WMI;

namespace win_int_info.Serialization
{
    internal sealed class DiskPropMap : CsvClassMap<DiskProperties>
    {
        public DiskPropMap()
        {
            this.AutoMap();

        }
    }
}