using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Win32;

namespace DCOM.Global
{
    /// <summary>
    /// Class for handling Windows Registry actions
    /// </summary>
    public static class RegistryFunctions
    {
         
        /// <summary>
        /// Export a registry key using cmdline <code>reg export</code> and save as a .reg file on disk
        /// </summary>
        /// <param name="keyPath">Full path to registry key to be exported. All child registry keys will also be exported. </param>
        /// <param name="outputFileName">Full path to .reg file to be created. Any existing file of the same name will be overwritten</param>
        /// <returns></returns>
        /// <example>
        /// <code>ExportRegistryKey("HKLM\\SOFTWARE\\WOW6432Node\\ODBC\\ODBC.INI","C:\\temp\\odbc64.reg")</code>
        /// </example>
        public static bool ExportRegistryKey(string keyPath, string outputFileName)
        {
            Process process = new Process();
            ProcessStartInfo sInfo = new ProcessStartInfo();

            List<string> stdMessages = new List<string>();
            List<string> stdErrors = new List<string>();

            process.EnableRaisingEvents = true;
            process.OutputDataReceived += new DataReceivedEventHandler((sender, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    stdMessages.Add(e.Data);
                }
            });
            process.ErrorDataReceived += new DataReceivedEventHandler((sender, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    stdErrors.Add(e.Data);
                }
            });
            sInfo.FileName = "cmd";

            sInfo.Arguments = string.Format(@"/c reg export ""{0}"" ""{1}"" /y ", keyPath, outputFileName);
            sInfo.RedirectStandardError = true;
            sInfo.RedirectStandardOutput = true;
            sInfo.UseShellExecute = false;
            sInfo.WindowStyle = ProcessWindowStyle.Hidden;
            process.StartInfo = sInfo;
            process.Start();
            if (process.StartInfo.RedirectStandardError)
            {
                process.BeginErrorReadLine();
            }

            if (process.StartInfo.RedirectStandardOutput)
            {
                process.BeginOutputReadLine();
            }

            process.WaitForExit();

            if (stdMessages.Count < 1 && stdMessages[stdMessages.Count - 1]
                .StartsWith("The operation completed successfully", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (stdErrors.Count >= 1)
            {
                var lst = stdErrors;
                throw new Exception(stdErrors[stdErrors.Count - 1]);
            }

            return false;
        }

        /// <summary>
        /// Import a registry key from a .reg file on disk. .reg must have been created using cmdline <code>reg export</code> or <seealso cref="ExportRegistryKey"/>
        /// </summary>
        /// <param name="inputFileName">Full path to .reg file to be imported</param>
        /// <returns></returns>
        public static bool ImportRegistryKey(string inputFileName)
        {
            Process process = new Process();
            ProcessStartInfo sInfo = new ProcessStartInfo();
            List<string> stdMessages = new List<string>();
            List<string> stdErrors = new List<string>();
            process.EnableRaisingEvents = true;
            process.OutputDataReceived += new DataReceivedEventHandler((sender, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    stdMessages.Add(e.Data);
                }
            });
            process.ErrorDataReceived += new DataReceivedEventHandler((sender, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    stdErrors.Add(e.Data);
                }
            });
            sInfo.FileName = "cmd";

            sInfo.Arguments = string.Format(@"/c reg import ""{0}"" ", inputFileName);
            sInfo.RedirectStandardError = true;
            sInfo.RedirectStandardOutput = true;
            sInfo.UseShellExecute = false;
            sInfo.WindowStyle = ProcessWindowStyle.Hidden;
            process.StartInfo = sInfo;
            process.Start();
            if (process.StartInfo.RedirectStandardError)
            {
                process.BeginErrorReadLine();
            }

            if (process.StartInfo.RedirectStandardOutput)
            {
                process.BeginOutputReadLine();
            }

            process.WaitForExit();

            if (stdMessages.Count < 1 && stdMessages[stdMessages.Count - 1]
                .StartsWith("The operation completed successfully", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (stdErrors.Count >= 1)
            {
                throw new Exception(stdErrors[stdErrors.Count - 1]);
            }

            return false;
        }

        /// <summary>
        /// Return registry key
        /// </summary>
        /// <param name="path"></param>
        /// <param name="computerName"></param>
        /// <param name="hive"></param>
        /// <returns></returns>
        internal static RegistryKey GetKey(string path, string computerName = null, RegistryHive hive = RegistryHive.LocalMachine)
        {
            RegistryKey key;
            if (!string.IsNullOrEmpty(computerName))
            {
                key = RegistryKey.OpenRemoteBaseKey(hive, computerName);
            }
            else
            {
                switch (hive)
                {
                    case RegistryHive.ClassesRoot:
                        key = Registry.ClassesRoot;
                        break;
                    case RegistryHive.CurrentConfig:
                        key = Registry.CurrentConfig;
                        break;
                    case RegistryHive.CurrentUser:
                        key = Registry.CurrentUser;
                        break;
                    case RegistryHive.DynData:
                        key = Registry.DynData;
                        break;
                    case RegistryHive.PerformanceData:
                        key = Registry.PerformanceData;
                        break;
                    case RegistryHive.Users:
                        key = Registry.Users;
                        break;
                    default:
                        key = Registry.LocalMachine;
                        break;
                }
            }
            var subKey = key.OpenSubKey(path, false);
            if (subKey == null)
                throw new InvalidOperationException(
                    $"{computerName} : The registry key '{hive}:\\{path}' was not found.");
            return subKey;
        }

        internal static Dictionary<RegistryKey, string> GetAllChildren(RegistryKey key, string keyPath)
        {
            var rtList = new Dictionary<RegistryKey, string>();
            
            foreach (var name in key.GetSubKeyNames())
            {
                var path = $"{keyPath}\\{name}";
                var curKey = key.OpenSubKey(name, false);
                if (curKey == null) return rtList;
                rtList.Add(curKey, path);
                if(curKey.GetSubKeyNames().Any())
                {
                    GetAllChildren(curKey, path).ToList().ForEach(x => rtList.Add(x.Key, x.Value));
                }
            }

            return rtList;
        }
    }
}
