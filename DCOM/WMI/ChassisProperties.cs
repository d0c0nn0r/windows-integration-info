using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Net;
using Newtonsoft.Json;

namespace WinDevOps.WMI
{
    /// <summary>
    /// Chassis specific properties
    /// </summary>
    /// <remarks>
    /// <see href="https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-systemenclosure"/>
    /// </remarks>
    public class ChassisProperties
    {
        [JsonIgnore]
        private static readonly List<string> ChassisModels = new List<string>()
        {
            "PlaceHolder", "Maybe Virtual Machine", "Unknown", "Desktop", "Thin Desktop", "Pizza Box", "Mini Tower",
            "Full Tower", "Portable",
            "Laptop", "Notebook", "Hand Held", "Docking Station", "All in One", "Sub Notebook", "Space-Saving",
            "Lunch Box", "Main System Chassis",
            "Lunch Box", "SubChassis", "Bus Expansion Chassis", "Peripheral Chassis", "Storage Chassis",
            "Rack Mount Unit", "Sealed-Case PC"
        };

        public string ChassisModel
        {
            get
            {
                return ChassisModels[(int) this.ChassisTypes[0]];
            }
        }

        public UInt16[] ChassisTypes { get; internal set; }
        public string Manufacturer { get; internal set; }
        public string SerialNumber { get; internal set; }
        public string Tag { get; internal set; }
        public string SKU { get; internal set; }
       
        public Dictionary<string, object> Miscellaneous { get; internal set; }
        /// <summary>
        /// Default internal constructor
        /// </summary>
        internal ChassisProperties()
        {
            Miscellaneous = new Dictionary<string, object>();
        }
        internal static ChassisProperties Load(ManagementScope scope)
        {
            var biosCollection = WMIUtils.PerformQuery(scope, "SELECT * FROM Win32_SystemEnclosure")?.ToList();
            if (biosCollection == null) return null;

            if (!WMIUtils.Initialize<ChassisProperties>(biosCollection.First(), out var x))
            {
                throw new Exception("Unable to serialize WMI Chassis Properties");
            }
            return x;
        }

        public static ChassisProperties Load(string computerName = null,
            NetworkCredential credential = null)
        {
            ManagementScope scope = WMIUtils.GetScope(computerName, credential);
            return Load(scope);
        }

    }
}