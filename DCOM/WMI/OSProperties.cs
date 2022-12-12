using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Net;
using Newtonsoft.Json;

namespace WinDevOps.WMI
{
    /// <summary>
    /// Operating System specific properties
    /// </summary>
    /// <remarks>
    /// <see href="https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-operatingsystem"/>
    /// </remarks>
    public class OSProperties
    {
        public string BuildNumber { get; internal set; }
        public string Version { get; internal set; }
        public string SerialNumber { get; internal set; }
        public UInt16 ServicePackMajorVersion { get; internal set; }
        public object CSDVersion { get; internal set; }
        public string SystemDrive { get; internal set; }
        public string SystemDirectory { get; internal set; }
        public string WindowsDirectory { get; internal set; }
        public string Caption { get; internal set; }
        public UInt64 TotalVisibleMemorySize { get; internal set; }
        public UInt64 FreePhysicalMemory { get; internal set; }
        public UInt64 TotalVirtualMemorySize { get; internal set; }
        public UInt64 FreeVirtualMemory { get; internal set; }
        public string OSArchitecture { get; internal set; }
        public string Organization { get; internal set; }
        public string LocalDateTime { get; internal set; }
        public string RegisteredUser { get; internal set; }
        public UInt32 OperatingSystemSKU { get; internal set; }
        public UInt16 OSType { get; internal set; }
        public string LastBootUpTime { get; internal set; }
        public string InstallDate { get; internal set; }
        public bool IsVirtual { get; internal set; }
        public string VirtualType { get; internal set; }

        public Dictionary<string, object> Miscellaneous { get; internal set; }

        [JsonIgnore]
        private static readonly List<string> _SKUs = new List<string>
        {
            "Undefined", "Ultimate Edition", "Home Basic Edition", "Home Basic Premium Edition", "Enterprise Edition",
            "Home Basic N Edition", "Business Edition", "Standard Server Edition", "DatacenterServer Edition",
            "Small Business Server Edition",
            "Enterprise Server Edition", "Starter Edition", "Datacenter Server Core Edition",
            "Standard Server Core Edition",
            "Enterprise ServerCoreEdition", "Enterprise Server Edition for Itanium-Based Systems", "Business N Edition",
            "Web Server Edition",
            "Cluster Server Edition", "Home Server Edition", "Storage Express Server Edition",
            "Storage Standard Server Edition",
            "Storage Workgroup Server Edition", "Storage Enterprise Server Edition",
            "Server For Small Business Edition", "Small Business Server Premium Edition"
        };

        public string OSSku
        {
            get;
            internal set;
        }
        /// <summary>
        /// Default internal constructor
        /// </summary>
        internal OSProperties()
        {
            Miscellaneous = new Dictionary<string, object>();
        }

        [JsonIgnore]
        private readonly ManagementObject _wmiBaseObject;

        /// <summary>
        /// Default internal constructor
        /// </summary>
        internal OSProperties(ManagementObject wmiObject)
        {
            Miscellaneous = new Dictionary<string, object>();
            _wmiBaseObject = wmiObject;
            if (_wmiBaseObject.GetPropertyValue("OperatingSystemSKU") != null)
            {
                OSSku= _SKUs[(int)this.OperatingSystemSKU];
            }
            OSSku ="Not Available";
        }

        internal static OSProperties Load(ManagementScope scope)
        {
            var biosCollection = WMIUtils.PerformQuery(scope, "SELECT * FROM Win32_OperatingSystem")?.ToList();
            var compCollection = WMIUtils.PerformQuery(scope, "SELECT Manufacturer, Model FROM Win32_ComputerSystem")?.ToList();
            if (biosCollection == null) return null;
            if (compCollection == null) return null;

            if (!WMIUtils.Initialize<OSProperties>(biosCollection.First(), out var x))
            {
                throw new Exception("Unable to serialize WMI OS Properties");
            }
            x.IsVirtual =
                WMIUtils.ResolveVirtualMachine(biosCollection.First(), compCollection.First(), out var virtualType);
            x.VirtualType = virtualType;
            return x;
        }

        public static OSProperties Load(string computerName = null,
            NetworkCredential credential = null)
        {
            ManagementScope scope = WMIUtils.GetScope(computerName, credential);
            return Load(scope);
        }
    }
}