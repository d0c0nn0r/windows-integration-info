using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Net;

namespace WinDevOps.WMI
{
    /// <summary>
    /// Processor specific properties
    /// </summary>
    /// <remarks>
    /// <see href="https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-processor"/>
    /// </remarks>
    public class ProcessorProperties
    {
        public string Name { get; internal set; }
        public string Description { get; internal set; }
        public UInt32 MaxClockSpeed { get; internal set; }
        public UInt32 CurrentClockSpeed { get; internal set; }
        public UInt16 AddressWidth { get; internal set; }
        public UInt32 NumberOfCores { get; internal set; }
        public UInt32 NumberOfLogicalProcessors { get; internal set; }
        public UInt16 DataWidth { get; internal set; }
        
        public Dictionary<string, object> Miscellaneous { get; internal set; }
        /// <summary>
        /// Default internal constructor
        /// </summary>
        internal ProcessorProperties()
        {
            Miscellaneous = new Dictionary<string, object>();
        }
        internal static ProcessorProperties Load(ManagementScope scope)
        {
            var biosCollection = WMIUtils.PerformQuery(scope, "SELECT * FROM Win32_Processor")?.ToList();
            if (biosCollection == null) return null;

            if (!WMIUtils.Initialize<ProcessorProperties>(biosCollection.First(), out var x))
            {
                throw new Exception("Unable to serialize WMI Processor Properties");
            }
            return x;
        }

        public static ProcessorProperties Load(string computerName = null,
            NetworkCredential credential = null)
        {
            ManagementScope scope = WMIUtils.GetScope(computerName, credential);
            return Load(scope);
        }
    }
}