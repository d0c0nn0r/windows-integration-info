using CsvHelper.Configuration;
using WinDevOps.Accounts;

namespace win_int_info.Serialization
{
    internal sealed class WindowsUserMap : CsvClassMap<WindowsUser>
    {
        public WindowsUserMap()
        {
            this.AutoMap();

        }
    }
}