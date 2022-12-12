using CsvHelper.Configuration;
using WinDevOps;

namespace win_int_info.Serialization
{
    internal sealed class DcomGblMap : CsvClassMap<DCOMGlobal>
    {
        public DcomGblMap()
        {
            this.AutoMap();
            this.Map(m => m.ComputerName).Name("PSComputerName");
            
            this.Map(m => m.DefaultAccessPermissions).Name("DefaultAccessPermissions").TypeConverter<EnumerableConverter>();
            this.Map(m => m.LimitsAccessPermissions).Name("LimitsAccessPermissions").TypeConverter<EnumerableConverter>();
            this.Map(m => m.DefaultLaunchAndActivation).Name("DefaultLaunchAndActivation").TypeConverter<EnumerableConverter>();
            this.Map(m => m.LimitsLaunchAndActivation).Name("LimitsLaunchAndActivation").TypeConverter<EnumerableConverter>();
            this.Map(m => m.DcomProtocols).Name("DcomProtocols").TypeConverter<EnumerableConverter>();

        }
    }
}