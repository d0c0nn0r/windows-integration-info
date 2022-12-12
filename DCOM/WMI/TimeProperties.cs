using DCOM.Global;
using Microsoft.Win32;

namespace WinDevOps.WMI
{
    /// <summary>
    /// Time synchronization properties
    /// </summary>
    public class TimeProperties
    {
        public bool? DSTEnabled { get; internal set; }
        public string TimeZone { get; internal set; }
        public string ShortDate { get; internal set; }
        public string TimeFormat { get; internal set; }

        internal static TimeProperties Load(string computerName = null)
        {
            using (var regionSettings = RegistryFunctions.GetKey("Control Panel\\International", computerName, RegistryHive.CurrentUser))
            {
                var tzInfo = RegistryFunctions.GetKey("SYSTEM\\CurrentControlSet\\Control\\TimeZoneInformation", computerName,
                    RegistryHive.LocalMachine);
                bool? dst = null;
                try
                {
                    dst = ((int) tzInfo.GetValue("DynamicDaylightTimeDisabled")) == 0;
                }
                catch
                {
                    dst = null;
                }

                var r = new TimeProperties()
                {
                    DSTEnabled = dst,
                    TimeZone = tzInfo.GetValue("TimeZoneKeyName")?.ToString(),
                    ShortDate = regionSettings.GetValue("sShortDate")?.ToString(),
                    TimeFormat = regionSettings.GetValue("sTimeFormat")?.ToString()
                };
                tzInfo.Close();
                return r;
            }
        }
    }
}