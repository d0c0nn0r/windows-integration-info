using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using DCOM.Global;
using Microsoft.Win32;

namespace DCOM.Software_Management
{
    public static class SoftwareUtils
    {
        private static readonly List<string> RegistryLocations = new List<string>(2){"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\",
            "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\"};

        private static readonly string DotNetLocation = "SOFTWARE\\Microsoft\\NET Framework Setup\\NDP";
        
        /// <summary>
        /// Retrieve all installed software on the target computer
        /// </summary>
        /// <param name="computerName">Name of the computer to retrieve software from. If blank, local computer is assumed</param>
        /// <returns></returns>
        public static List<SoftwareProgram> GetSoftware(string computerName = null)
        {
            var returnList = new List<SoftwareProgram>();
            var parentKeys = RegistryLocations
                .Select(item => RegistryFunctions.GetKey(item, computerName, RegistryHive.LocalMachine))
                .ToList();
            try
            {
                //all software that would appear in Control Panel
                foreach (var regKey in parentKeys)
                {
                    foreach (var sName in regKey.GetSubKeyNames())
                    {

                        var curKey = regKey.OpenSubKey(sName, false);
                        if(curKey == null) continue;
                        var displayName = curKey.GetValue("DisplayName")
                            ?.ToString()
                            .Trim()
                            .Replace("  "," ");
                        if(string.IsNullOrEmpty(displayName)) continue; //skip reg keys that have no displayName

                        var version = curKey.GetValue("DisplayVersion")?.ToString();
                        var publisher = curKey.GetValue("Publisher")?.ToString();
                        var software = new SoftwareProgram()
                        {
                            ComputerName = computerName ?? Environment.MachineName, 
                            ProgramName = displayName, 
                            Publisher = publisher,
                            Version = version,
                            RegistryKey = curKey.Name
                        };

                        var dontMisc = new List<string>(3) {"DisplayName", "DisplayVersion", "Publisher"};
                        var misc = curKey.GetValueNames()
                            .Where(name => !dontMisc.Contains(name, StringComparer.OrdinalIgnoreCase));
                        foreach (var miscName in misc)
                        {
                            software.Miscellaneous.Add(miscName, curKey.GetValue(miscName));
                        }

                        returnList.Add(software);
                    }
                }

                var dotNets = GetDotNetFrameworks(computerName);
                if (dotNets.Any()) returnList.AddRange(dotNets);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"{computerName} : Error retrieving software. {ex.Message}");
            }
            return returnList;
        }

        /// <summary>
        /// Retrieve DotNet Framework versions installed
        /// </summary>
        /// <param name="computerName"></param>
        /// <returns></returns>
        public static List<SoftwareProgram> GetDotNetFrameworks(string computerName = null)
        {
            var returnList = new List<SoftwareProgram>();
            var rootKey = RegistryFunctions.GetKey(DotNetLocation, computerName);
            // get only keys where Name does not start with an S : (?!S)
            // and must start with a letter character \p{L}
            var regex = new Regex("^(?!S)\\p{L}", RegexOptions.None);
            var dotNetKeys = RegistryFunctions.GetAllChildren(rootKey, "SOFTWARE\\Microsoft\\NET Framework Setup\\NDP")
                .Where(k => k.Key.GetValueNames().Contains("Version"))?.ToList();
            
            foreach (var dotNet in dotNetKeys?.Where(key => regex.IsMatch(key.Key.Name.Split('\\').Last())))
            {
                var product = ResolveFrameworkVersion(dotNet.Key.GetValue("Release"));
                var keyPath = dotNet.Value.Split('\\');
                var match = new Regex("(?i)(v((\\d+\\.)([^\\s]+)|\\d+))", RegexOptions.IgnoreCase);
                var majorVersionLookup=0;
                for (var i = 0; i < keyPath.Length; i++)
                {
                    if (!match.IsMatch(keyPath[i])) continue;
                    majorVersionLookup = i;
                    break;
                }

                string programName;
                var spName = dotNet.Key.GetValue("SP")?.ToString();
                if (!string.IsNullOrEmpty(product))
                {
                    var str = string.Join(" ", (keyPath.Skip(majorVersionLookup + 1).Take(keyPath.Length)).ToArray());
                    str = Regex.Replace(str, "\\b(?i)s[a-zA-Z]*", "").Replace("  ", " ");
                    programName = string.IsNullOrEmpty(spName) ? $"Microsoft .NET Framework v{product} {str}" : $"Microsoft .NET Framework v{product} SP {spName} {str}";
                }
                else
                {
                    var str = string.Join(" ", (keyPath.Skip(majorVersionLookup).Take(keyPath.Length)).ToArray());
                    str = Regex.Replace(str, "\\b(?i)s[a-zA-Z]*", "").Replace("  ", " ");
                    programName = string.IsNullOrEmpty(spName) ? $"Microsoft .NET Framework {str}" : $"Microsoft .NET Framework {str} SP{spName}";
                }

                returnList.Add(new SoftwareProgram()
                {
                    ComputerName = computerName,
                    ProgramName = programName,
                    Publisher = "Microsoft Corporation",
                    Version = dotNet.Key.GetValue("Version")?.ToString()
                });
            }

            return returnList;
        }

        /// <summary>
        /// Resolve binary version to product version
        /// </summary>
        /// <param name="version"></param>
        /// <returns></returns>
        private static string ResolveFrameworkVersion(object version)
        {
            if(version == null) return String.Empty;
            switch (version)
            {
                case var _ when (int)version >= 528040: { return "4.8"; }
                case var _ when (int)version >= 461808:  { return "4.7.2"; }
                case var _ when (int)version >= 461308: { return "4.7.1"; }
                case var _ when (int)version >= 460798: { return "4.7"; }
                case var _ when (int)version >= 394802: { return "4.6.2"; }
                case var _ when (int)version >= 394254: { return "4.6.1"; }
                case var _ when (int)version >= 393295: { return "4.6"; }
                case var _ when (int)version >= 379893: { return "4.5.2"; }
                case var _ when (int)version >= 378675: { return "4.5.1"; }
                case 378389: { return "4.5"; }
                //case 378675: { return "4.5.1"; }
                //case 378758: { return "4.5.1"; }
                //case 379893: { return "4.5.2"; }
                //case 393295: { return "4.6"; }
                //case 393297: { return "4.6"; }
                //case 394254: { return "4.6.1"; }
                //case 394271: { return "4.6.1"; }
                //case 394802: { return "4.6.2"; }
                //case 394806: { return "4.6.2"; }
                //case 460798: { return "4.7"; }
                //case 460805: { return "4.7"; }
                //case 461308: { return "4.7.1"; }
                //case 461310: { return "4.7.1"; }
                //case 461808: { return "4.7.2"; }
                //case 461814: { return "4.7.2"; }
                //case 528040: { return "4.8"; }
                //case 528049: { return "4.8"; }
                //case 528209: { return "4.8"; }
                //case 528372: { return "4.8"; }
                //case 528449: { return "4.8"; }
            }

            return string.Empty;
        }
    }
}
