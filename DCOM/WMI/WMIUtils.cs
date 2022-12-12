using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Net;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace WinDevOps.WMI
{
    public static class WMIUtils
    {
        /// <summary>
        /// Format a number representing BYTES to it's nearest MB, GB, TB string equivalent
        /// </summary>
        /// <param name="value"></param>
        /// <param name="decimalPlaces"></param>
        /// <returns></returns>
        public static string FormatBytesToSize(Int64 value, int decimalPlaces = 1)
        {
            string[] SizeSuffixes =
                { "bytes", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB" };
            if (decimalPlaces < 0) { throw new ArgumentOutOfRangeException("decimalPlaces"); }
            if (value < 0) { return "-" + FormatBytesToSize(-value, decimalPlaces); }
            if (value == 0) { return string.Format("{0:n" + decimalPlaces + "} bytes", 0); }

            // mag is 0 for bytes, 1 for KB, 2, for MB, etc.
            int mag = (int)Math.Log(value, 1024);

            // 1L << (mag * 10) == 2 ^ (10 * mag) 
            // [i.e. the number of bytes in the unit corresponding to mag]
            decimal adjustedSize = (decimal)value / (1L << (mag * 10));

            // make adjustment when the value is large enough that
            // it would round up to 1000 or more
            if (Math.Round(adjustedSize, decimalPlaces) >= 1000)
            {
                mag += 1;
                adjustedSize /= 1024;
            }

            return string.Format("{0:n" + decimalPlaces + "} {1}",
                adjustedSize,
                SizeSuffixes[mag]);
        }
        /// <summary>
        /// Gather all hardware specific details, including <see cref="DiskProperties"/>, <see cref="OSProperties"/>, <see cref="NetworkCardProperties"/>, <see cref="TimeProperties"/>, and <see cref="ComputerProperties"/>
        /// </summary>
        /// <param name="computerName">Name of remote computer to gather details for. Leave null or blank for local computer</param>
        /// <param name="credential">Credentials to use when authenticating to a remote computer. Leave null or blank to execute under context of current user</param>
        /// <returns></returns>
        public static ComputerInformation GetComputerInformation(string computerName=null, NetworkCredential credential = null)
        {
            ComputerInformation info = new ComputerInformation();
            var connectOpts = new ConnectionOptions();
            if (credential != null)
            {
                connectOpts.Username = string.IsNullOrEmpty(credential.Domain)
                    ? $"{credential.UserName}"
                    : $"{credential.Domain}\\{credential.UserName}";
                connectOpts.Password = credential.Password;
            }
            ManagementScope scope = GetScope(computerName, credential);
            
            #region ComputerProps

            try
            {
                info.ComputerSystem = ComputerProperties.Load(scope);
            }
            catch (Exception ex)
            {
                //failed to get network card details
                Console.Error.WriteLine("Failed to gather Computer System details", ex);
            }
            #endregion

            #region Bios
            try
            {
                info.Bios = BIOSProperties.Load(scope);
            }
            catch (Exception ex)
            {
                //failed to get network card details
                Console.Error.WriteLine("Failed to gather BIOS details", ex);
            }
            #endregion

            #region OS
            try
            {
                info.OperatingSystem = OSProperties.Load(scope);
            }
            catch (Exception ex)
            {
                //failed to get network card details
                Console.Error.WriteLine("Failed to gather OS details", ex);
            }
            #endregion

            #region Chassis
            try
            {
                info.Chassis = ChassisProperties.Load(scope);
            }
            catch (Exception ex)
            {
                //failed to get network card details
                Console.Error.WriteLine("Failed to gather Chassis details", ex);
            }
            #endregion

            #region Processor
            try
            {
                info.ProcessorProperties = ProcessorProperties.Load(scope);
            }
            catch (Exception ex)
            {
                //failed to get network card details
                Console.Error.WriteLine("Failed to gather Processor details", ex);
            }
            #endregion


            #region Memory
            try
            {
                var memoryCollection = MemoryProperties.Load(scope);
                var arrayCollection = MemoryArrayProperties.Load(scope);
                info.MemoryProperties.AddRange(memoryCollection);
                info.MemoryArrayProperties.AddRange(arrayCollection);
            }
            catch (Exception ex)
            {
                //failed to get network card details
                Console.Error.WriteLine("Failed to gather Memory details", ex);
            }
            #endregion

            #region NetworkCards

            try
            {
                var cards = NetworkCardProperties.Load(scope, computerName, credential);
                if (cards.Any()) info.NetworkCards.AddRange(cards);
            }
            catch(Exception ex)
            {
                //failed to get network card details
                Console.Error.WriteLine("Failed to gather network card details", ex);
            }
            #endregion

            #region Disks
            try
            {
                var disks = DiskProperties.Load(scope);
                info.Disks.AddRange(disks);
            }
            catch (Exception ex)
            {
                //failed to get network card details
                Console.Error.WriteLine("Failed to gather Disk details", ex);
            }
            #endregion Disks

            #region Time
            try
            {
                info.Time = TimeProperties.Load(computerName);
            }
            catch (Exception ex)
            {
                //failed to get network card details
                Console.Error.WriteLine("Failed to gather TIME details", ex);
            }
            #endregion Time
            return info;
        }

        /// <summary>
        /// Reflection to populate class fields dynamically from WMI object
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="baseObject"></param>
        /// <param name="returnObject"></param>
        /// <returns></returns>
        /// <remarks>
        /// Reflection used to cut down on boilerplate code in classes.
        /// </remarks>
        internal static bool Initialize<T>(ManagementObject baseObject, out T returnObject)
        {
            Type[] emptyArgs = Type.EmptyTypes;
            var constructorInfo = typeof(T).GetConstructor(BindingFlags.Public| BindingFlags.NonPublic | BindingFlags.Instance, null, emptyArgs, null);
            var constructorInfo2 = typeof(T).GetConstructor(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new Type[]{typeof(ManagementObject)}, null);
            if (constructorInfo == null)
                throw new NotSupportedException(
                    $"{typeof(T)} must have a default constructor with no arguments.");
            if(constructorInfo2 !=null)
            {
                returnObject = (T)constructorInfo2.Invoke(new object[] { baseObject });
            }
            else
            {
                returnObject = (T)constructorInfo.Invoke(new object[] { });
            }
            var compProps = typeof(T).GetProperties().Select(p => p.Name).ToList();
            var objProps = baseObject.Properties.OfType<PropertyData>().Select(p => p.Name).ToList();
            foreach (var prop in compProps)
            {
                if(objProps.Contains(prop, StringComparer.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"{prop} = {baseObject.GetPropertyValue(prop)}");
                    typeof(T).GetProperty(prop)
                        .SetValue(returnObject, baseObject.GetPropertyValue(prop), null);
                }
            }

            try
            {
                var miscProp = typeof(T).GetProperty("Miscellaneous");
                var method = miscProp.PropertyType
                    .GetMethod("Add", new[] {typeof(string), typeof(object)});
                foreach (var prop in baseObject.Properties)
                {
                    if (method != null && !compProps.Contains(prop.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        method.Invoke(miscProp.GetValue(returnObject,null), new object[] {prop.Name, prop.Value});
                    }
                }
            }
            catch
            {
                //miscellaneous properties can be ignored as 'non-essential' and therefore we don't necessarily care about exceptions here
            }

            return true;
        }

        /// <summary>
        /// Create Scope object
        /// </summary>
        /// <param name="computerName"></param>
        /// <param name="credential"></param>
        /// <param name="nameSpace"></param>
        /// <returns></returns>
        internal static ManagementScope GetScope(string computerName = null, NetworkCredential credential = null, string nameSpace = "\\root\\cimv2")
        {
            var connectOpts = new ConnectionOptions();
            if (credential != null)
            {
                connectOpts.Username = string.IsNullOrEmpty(credential.Domain)
                    ? $"{credential.UserName}"
                    : $"{credential.Domain}\\{credential.UserName}";
                connectOpts.Password = credential.Password;
            }

            ManagementScope scope;
            if (string.IsNullOrEmpty(computerName))
            {
                scope = new ManagementScope(nameSpace, connectOpts);
            }
            else
            {
                scope = new ManagementScope($"\\\\{computerName}{nameSpace}", connectOpts);
            }
            scope.Connect();
            if (!scope.IsConnected)
                throw new ManagementException(
                    $"Failed to connect to the application server.");
            return scope;
        }
        /// <summary>
        /// Execute WMI query
        /// </summary>
        /// <param name="scope"></param>
        /// <param name="query"></param>
        /// <returns></returns>
        internal static IEnumerable<ManagementObject> PerformQuery(ManagementScope scope, string query)
        {
            ObjectQuery q = new ObjectQuery(query);
            using (var searcher = new ManagementObjectSearcher(scope, q))
            {
                return searcher.Get().OfType<ManagementObject>();
            }
        }

        /// <summary>
        /// Resolve virtual machine type from WMI information
        /// </summary>
        /// <param name="bios"></param>
        /// <param name="comp"></param>
        /// <param name="virtualType"></param>
        /// <returns></returns>
        internal static bool ResolveVirtualMachine(ManagementObject bios, ManagementObject comp, out string virtualType)
        {
            virtualType = string.Empty;
            if (Regex.IsMatch(bios.GetPropertyValue("Version").ToString(), "VIRTUAL"))
            {
                virtualType = "Virtual - Hyper-V";
                return true;
            }
            else if (Regex.IsMatch(bios.GetPropertyValue("Version").ToString(), "A M I"))
            {
                virtualType = "Virtual - Virtual PC";
                return true;
            }
            else if (bios.GetPropertyValue("Version").ToString().Contains("Xen"))
            {
                virtualType = "Virtual - Xen";
                return true;
            }
            else if (bios.GetPropertyValue("Version").ToString().Contains("VMware"))
            {
                virtualType = "Virtual - VMWare";
                return true;
            }
            else if (comp.GetPropertyValue("manufacturer").ToString().Contains("Microsoft"))
            {
                virtualType = "Virtual - Hyper-V";
                return true;
            }
            else if (comp.GetPropertyValue("manufacturer").ToString().Contains("VMWare"))
            {
                virtualType = "Virtual - VMWare";
                return true;
            }
            else if (comp.GetPropertyValue("model").ToString().Contains("Virtual"))
            {
                virtualType = "Unknown Virtual Machine";
                return true;
            }

            return false;
        }
    }
}
