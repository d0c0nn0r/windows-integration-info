using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Net;
using Newtonsoft.Json;

namespace WinDevOps.WMI
{
    /// <summary>
    /// Network card properties
    /// </summary>
    /// <remarks>
    /// <see href="https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-networkadapter"/></remarks>
    public class NetworkCardProperties
    {
        public string AdapterName { get; internal set; }
        public Net_Connection_Status ConnectionStatus { get; internal set; }
        public object DefaultIPGateway { get; internal set; }
        public string Description { get; internal set; }
        public bool DHCPEnabled { get; internal set; }
        public string DHCPServer { get; internal set; }
        public object DNSDomain { get; internal set; }
        public string[] DNSDomainSuffixSearchOrder { get; internal set; }
        public object DNSServerSearchOrder { get; internal set; }
        public bool DomainDNSRegistrationEnabled { get; internal set; }
        public UInt32 Index { get; internal set; }
        public UInt32 InterfaceIndex { get; internal set; }
        public string[] IpAddress { get; internal set; }
        public string[] IpSubnet { get; internal set; }
        public string MACAddress { get; internal set; }
        public string NetConnectionId { get; internal set; }
        public string NetworkName { get; internal set; }
        /// <summary>
        /// 
        /// </summary>
        /// <remarks>
        /// <see href="https://en.wikipedia.org/wiki/Promiscuous_mode"/>
        /// </remarks>
        public bool PromiscuousMode => (Modes & NDIS_PACKET_TYPE.PROMISCUOUS) == NDIS_PACKET_TYPE.PROMISCUOUS;
        
        public object WinsPrimaryServer { get; internal set; }
        public object WinsSecondaryServer { get; internal set; }

        public NDIS_PACKET_TYPE Modes { get; internal set; }

        [JsonIgnore]
        private readonly ManagementObject _adapterWmiManagement;
        [JsonIgnore]
        private readonly ManagementObject _configWmiManagement;
        [JsonIgnore]
        private readonly ManagementObject _promiscWmiManagement;

        internal NetworkCardProperties(ManagementObject adapterWmi, ManagementObject configWmi, ManagementObject promiscWmi)
        {
            _adapterWmiManagement = adapterWmi;
            _configWmiManagement = configWmi;
            _promiscWmiManagement = promiscWmi;
            var pNames = _configWmiManagement.Properties.OfType<PropertyData>().Select(p => p.Name).ToList();
            
            this.AdapterName = _adapterWmiManagement.GetPropertyValue("Name").ToString();
            this.ConnectionStatus = _adapterWmiManagement.GetPropertyValue("NetConnectionStatus") == null ? Net_Connection_Status.Disconnected : (
                Net_Connection_Status)((UInt16)_adapterWmiManagement.GetPropertyValue("NetConnectionStatus"));
            this.NetConnectionId = _adapterWmiManagement.GetPropertyValue("NetConnectionId")?.ToString();
            this.Index = (uint)_configWmiManagement.GetPropertyValue("Index");
            
            ResolveIpAddress();
            
            this.MACAddress = _configWmiManagement.GetPropertyValue("MACAddress")?.ToString();
            this.DefaultIPGateway = pNames.Contains("DefaultIPGateway", StringComparer.OrdinalIgnoreCase) ? _configWmiManagement.GetPropertyValue("DefaultIPGateway") : null;
            this.InterfaceIndex = (uint)_configWmiManagement.GetPropertyValue("InterfaceIndex");
            this.Description = _configWmiManagement.GetPropertyValue("Description")?.ToString();
            this.DHCPEnabled = pNames.Contains("DHCPEnabled", StringComparer.OrdinalIgnoreCase) && (bool)_configWmiManagement.GetPropertyValue("DHCPEnabled");
            this.DHCPServer = pNames.Contains("DHCPServer", StringComparer.OrdinalIgnoreCase) ? _configWmiManagement.GetPropertyValue("DHCPServer")?.ToString() : null;
            this.DNSDomain = pNames.Contains("DNSDomain", StringComparer.OrdinalIgnoreCase) ? _configWmiManagement.GetPropertyValue("DNSDomain")?.ToString() : null;
            this.DNSDomainSuffixSearchOrder = pNames.Contains("DNSDomainSuffixSearchOrder", StringComparer.OrdinalIgnoreCase) ? (string[])_configWmiManagement.GetPropertyValue("DNSDomainSuffixSearchOrder") : null;
            this.DNSServerSearchOrder = pNames.Contains("DNSServerSearchOrder", StringComparer.OrdinalIgnoreCase) ? _configWmiManagement.GetPropertyValue("DNSServerSearchOrder")?.ToString() : null;
            this.DomainDNSRegistrationEnabled = pNames.Contains("DomainDNSRegistrationEnabled", StringComparer.OrdinalIgnoreCase) &&
                                                (_configWmiManagement.GetPropertyValue("DomainDNSRegistrationEnabled") != null && (bool)_configWmiManagement.GetPropertyValue(
                                                    "DomainDNSRegistrationEnabled"));
            this.WinsPrimaryServer = pNames.Contains("WinsPrimaryServer", StringComparer.OrdinalIgnoreCase) ? _configWmiManagement.GetPropertyValue("WinsPrimaryServer") : null;
            this.WinsSecondaryServer = pNames.Contains("WinsSecondaryServer", StringComparer.OrdinalIgnoreCase) ? _configWmiManagement.GetPropertyValue("WinsSecondaryServer") : null;
            ResolveModes();
            this.NetworkName = _adapterWmiManagement.GetPropertyValue("NetConnectionId")?.ToString();
        }

        internal static List<NetworkCardProperties> Load(ManagementScope scope, string computerName = null, NetworkCredential credential = null)
        {
            var returnList = new List<NetworkCardProperties>();
            
            var networkCollection = WMIUtils.PerformQuery(scope, "SELECT * FROM Win32_NetworkAdapter")?.ToList();
            if (networkCollection == null) throw new Exception("Unable to access WMI of target computer.");

            foreach (var adapter in networkCollection)
            {
                var netConfigQuery =
                    $"SELECT * FROM Win32_NetworkAdapterConfiguration WHERE Index = '{adapter.GetPropertyValue("Index")}'";
                var promiscQuery = $"SELECT * FROM MSNdis_CurrentPacketFilter WHERE InstanceName = '{adapter.GetPropertyValue("Name")}'";
                var promScope = WMIUtils.GetScope(computerName, credential, "\\root\\WMI");
                var promiscResult = WMIUtils.PerformQuery(promScope, promiscQuery)?.ToList().FirstOrDefault();
                var netConfigResult = WMIUtils.PerformQuery(scope, netConfigQuery).First();
                var nc = new NetworkCardProperties(adapter, netConfigResult, promiscResult);

                returnList.Add(nc);
            }

            return returnList;
        }

        public static List<NetworkCardProperties> Load(string computerName = null,
            NetworkCredential credential = null)
        {
            var scope = WMIUtils.GetScope(computerName, credential);
            return Load(scope, computerName, credential);
        }

        internal void ResolveModes()
        {
            //todo process each network card
            NDIS_PACKET_TYPE ncMode = NDIS_PACKET_TYPE.NONE;
            if (_promiscWmiManagement?.GetPropertyValue("NdisCurrentPacketFilter") != null)
            {
                var curX = _promiscWmiManagement?.GetPropertyValue("NdisCurrentPacketFilter");
                var modesList = Enum.GetValues(typeof(NDIS_PACKET_TYPE)).Cast<int>()
                    .Where(enumValue => enumValue != 0 && (enumValue & (UInt32)_promiscWmiManagement.GetPropertyValue("NdisCurrentPacketFilter")) == enumValue).ToList();
                foreach (var n in modesList)
                {
                    ncMode |= (NDIS_PACKET_TYPE)n; //add
                }

                ncMode &= ~NDIS_PACKET_TYPE.NONE;   //remove
            }

            this.Modes = ncMode;
        }
        internal void ResolveIpAddress()
        {
            string[] ip = null;
            string[] ipSubnet = null;
            try
            {
                ip = (string[])_configWmiManagement.GetPropertyValue("IpAddress");
            }
            catch
            {
                //not found
            }

            try
            {
                ipSubnet = (string[])_configWmiManagement.GetPropertyValue("IpSubnet");
            }
            catch
            {
                //not found
            }

            this.IpAddress = ip;
            this.IpSubnet = ipSubnet;
        }
    }
}