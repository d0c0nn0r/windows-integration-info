using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Net;
using Newtonsoft.Json;

namespace WinDevOps.WMI
{
    /// <summary>
    /// Memory card specific properties
    /// </summary>
    /// <remarks>
    /// <see href="https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-physicalmemory"/>
    /// </remarks>
    public class MemoryProperties
    {
        [JsonIgnore]
        private static readonly Dictionary<UInt16, string> MemoryDetail = new Dictionary<UInt16, string>()
        {
            {1, "Reserved"},
            {2, "Other"},
            {4, "Unknown"},
            {8, "Fast-paged"},
            {16, "Static column"},
            {32, "Pseudo-static"},
            {64, "RAMBUS"},
            {128, "Synchronous"},
            {256, "CMOS"},
            {512, "EDO"},
            {1024, "Window DRAM"},
            {2048, "Cache DRAM"},
            {4096, "Nonvolatile"}
        };
        [JsonIgnore]
        private static readonly List<string> MemoryModels = new List<string>()
        {
            "Unknown", "Other", "SIP", "DIP", "ZIP", "SOJ", "Proprietary", "SIMM", "DIMM", "TSOP", "PGA", "RIMM", 
            "SODIMM", "SRIMM", "SMD", "SSMP", "QFP", "TQFP", "SOIC", "LCC", "PLCC", "BGA", "FPBGA", "LGA"
        };
        

        public string BankLabel { get; internal set; }
        public string DeviceLocator { get; internal set; }

        public UInt64 Capacity { get; internal set; }
        public string CapacityString => WMIUtils.FormatBytesToSize((long)Capacity, 1);

        public string PartNumber { get; internal set; }
        public UInt32 Speed { get; internal set; }
        public string Tag { get; internal set; }
        public UInt16 FormFactor { get; internal set; }
        public string MemoryForm => MemoryModels[FormFactor];

        public UInt16 TypeDetail { get; internal set; }
        public string Detail => MemoryDetail[TypeDetail];

        public Dictionary<string, object> Miscellaneous { get; internal set; }
        /// <summary>
        /// Default internal constructor
        /// </summary>
        internal MemoryProperties()
        {
            Miscellaneous = new Dictionary<string, object>();
        }
        internal static List<MemoryProperties> Load(ManagementScope scope)
        {
            var lst = new List<MemoryProperties>();
            var biosCollection = WMIUtils.PerformQuery(scope, "SELECT * FROM Win32_PhysicalMemory")?.ToList();
            if (biosCollection == null) return null;
            foreach(var mem in biosCollection)
            {
                if (!WMIUtils.Initialize<MemoryProperties>(mem, out var x))
                {
                    throw new Exception("Unable to serialize WMI Memory Properties");
                }
                lst.Add(x);
            }
            return lst;
        }

        public static List<MemoryProperties> Load(string computerName = null,
            NetworkCredential credential = null)
        {
            ManagementScope scope = WMIUtils.GetScope(computerName, credential);
            return Load(scope);
        }
    }
}