using CsvHelper.Configuration;
using WinDevOps;

namespace win_int_info.Serialization
{
    internal sealed class DcomAppMap : CsvClassMap<DCOMApplication>
    {
        public DcomAppMap()
        {
            this.AutoMap();
            this.Map(m => m.ComputerName).Name("PSComputerName");
            this.Map(m => m.LaunchPermissions).Name("LaunchPermissions").TypeConverter<EnumerableConverter>();
            this.Map(m => m.AccessPermissions).Name("AccessPermissions").TypeConverter<EnumerableConverter>();
            this.Map(m => m.AuthenticationLevel).Name("AuthenticationLevel");

        }
    }
}