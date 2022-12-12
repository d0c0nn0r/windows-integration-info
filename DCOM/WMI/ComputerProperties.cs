using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Net;

namespace WinDevOps.WMI
{
    /// <summary>
    /// Generic computer properties
    /// </summary>
    /// <remarks>
    /// <see href="https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-computersystem"/>
    /// </remarks>
    public class ComputerProperties
    {
        public string DNSHostName { get; internal set; }
        public string Domain { get; internal set; }
        public string Manufacturer { get; internal set; }
        public string Model { get; internal set; }
        public UInt32 NumberOfLogicalProcessors { get; internal set; }
        public UInt32 NumberOfProcessors { get; internal set; }
        public string PrimaryOwnerContact { get; internal set; }
        public string PrimaryOwnerName { get; internal set; }
        public UInt64 TotalPhysicalMemory { get; internal set; }

        public string TotalPhysicalMemoryString => WMIUtils.FormatBytesToSize((long)TotalPhysicalMemory, 1);

        public Dictionary<string, object> Miscellaneous { get; internal set; }
        /// <summary>
        /// Default internal constructor
        /// </summary>
        internal ComputerProperties()
        {
            Miscellaneous = new Dictionary<string, object>();
        }


        internal static ComputerProperties Load(ManagementScope scope)
        {
            var biosCollection = WMIUtils.PerformQuery(scope, "SELECT * FROM Win32_ComputerSystem")?.ToList();
            if (biosCollection == null) return null;

            if (!WMIUtils.Initialize<ComputerProperties>(biosCollection.First(), out var x))
            {
                throw new Exception("Unable to serialize WMI OS Properties");
            }
            return x;
        }

        public static ComputerProperties Load(string computerName = null,
            NetworkCredential credential = null)
        {
            ManagementScope scope = WMIUtils.GetScope(computerName, credential);
            return Load(scope);
        }

    }
}