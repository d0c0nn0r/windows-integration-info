using CsvHelper.Configuration;
using WinDevOps.Accounts;

namespace win_int_info.Serialization
{
    internal sealed class WindowsGroupMap : CsvClassMap<WindowsGroup>
    {
        public WindowsGroupMap()
        {
            this.AutoMap();
            this.Map(m => m.Users).Name("Users").TypeConverter<EnumerableConverter>();

        }
    }
}