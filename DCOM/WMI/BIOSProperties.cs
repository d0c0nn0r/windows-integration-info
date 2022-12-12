using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Net;
using Newtonsoft.Json;

namespace WinDevOps.WMI
{
    /// <summary>
    /// BIOS specific properties
    /// </summary>
    /// <remarks>
    /// <see href="https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-bios"/>
    /// </remarks>
    public class BIOSProperties{
        public string Version { get; internal set; }
        public string SerialNumber { get; internal set; }
        
        public Dictionary<string, object> Miscellaneous { get; internal set; }
        

        /// <summary>
        /// Default internal constructor
        /// </summary>
        internal BIOSProperties()
        {
            Miscellaneous = new Dictionary<string, object>();
        }
        [JsonIgnore]
        private readonly ManagementObject _wmiBaseObject;

        /// <summary>
        /// Default internal constructor
        /// </summary>
        internal BIOSProperties(ManagementObject wmiObject) :this()
        {
            _wmiBaseObject = wmiObject;
        }

        internal static BIOSProperties Load(ManagementScope scope)
        {
            var biosCollection = WMIUtils.PerformQuery(scope, "SELECT * FROM Win32_BIOS")?.ToList();
            if (biosCollection == null) return null;

            if (!WMIUtils.Initialize<BIOSProperties>(biosCollection.First(), out var x))
            {
                throw new Exception("Unable to serialize WMI BIOS Properties");
            }
            return x;
        }

        public static BIOSProperties Load(string computerName = null,
            NetworkCredential credential = null)
        {
            ManagementScope scope = WMIUtils.GetScope(computerName, credential);
            return Load(scope);
        }
    }
}