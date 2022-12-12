using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Net;

namespace WinDevOps.WMI
{
    /// <summary>
    /// Memory Array information
    /// </summary>
    /// <remarks>
    /// <see href="https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-physicalmemoryarray"/>
    /// </remarks>
    public class MemoryArrayProperties
    {
        public UInt32 MaxCapacity { get; internal set; }
        public string Tag { get; internal set; }
        public UInt16 MemoryDevices { get; internal set; }
        
        public Dictionary<string, object> Miscellaneous { get; internal set; }
        /// <summary>
        /// Default internal constructor
        /// </summary>
        internal MemoryArrayProperties()
        {
            Miscellaneous = new Dictionary<string, object>();
        }

        internal static List<MemoryArrayProperties> Load(ManagementScope scope)
        {
            var lst = new List<MemoryArrayProperties>();
            var biosCollection = WMIUtils.PerformQuery(scope, "SELECT * FROM Win32_PhysicalMemoryArray")?.ToList();
            if (biosCollection == null) return null;
            foreach(var mem in biosCollection)
            {
                if (!WMIUtils.Initialize<MemoryArrayProperties>(mem, out var x))
                {
                    throw new Exception("Unable to serialize WMI OS Properties");
                }
                lst.Add(x);
            }
            return lst;
        }

        public static List<MemoryArrayProperties> Load(string computerName = null,
            NetworkCredential credential = null)
        {
            ManagementScope scope = WMIUtils.GetScope(computerName, credential);
            return Load(scope);
        }
    }
}