using System;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Security;

namespace WinDevOps.Process
{
    /// <summary>
    /// Process utilities
    /// </summary>
    public static class ProcessUtils
    {
        /// <summary>
        /// Invoke an application as a separate process, and optionally pass it arguments. Messages will be marshaled back to caller Console
        /// </summary>
        /// <param name="statements">Arguments to pass to the .exe</param>
        /// <param name="timeout">This is the number of seconds to allow process to run before timing out and killing it. Defaults to 300 seconds</param>
        /// <param name="fileName">The executable to call. Defaults to 'cmd.exe'.</param>
        /// <param name="workingDirectory">This sets the working directory that the executable or process is started in. Defaults to "C:\\windows\\system32"</param>
        /// <param name="noWindow">Run the process, without any window or console that can be interacted with</param>
        /// <param name="validExitCodes">Array of exit codes indicating success. Defaults to @(0).</param>
        /// <example>
        /// <para>Call <see href="https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/reg-export">reg export</see>
        /// and export the registry key where MS SQL Server information is stored.</para>
        /// <code lang="csharp">
        /// var str = string.Format("reg export \"{0}\" \"{1}\" /y ", (object)"HKLM\\SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server", 
        ///             (object) (Path.Combine("C:\\Temp", $"{Environment.MachineName}_Microsoft SQL Server_x64.reg")));
        /// var cmd = new string[]{"/c", str};
        /// WinDevOps.Process.ProcessUtils.InvokeProcess(cmd);
        /// </code>
        /// </example>
        public static void InvokeProcess(string[] statements, int timeout = 300, string fileName = "cmd.exe", 
            string workingDirectory = "C:\\windows\\system32", bool noWindow = true,
            int[] validExitCodes = null)
        {
            if (validExitCodes == null) validExitCodes = new[] {0};
            var startTime = DateTime.Now;
            var timespan = new TimeSpan(0, 0, 0, timeout);

            System.Diagnostics.Process process = new System.Diagnostics.Process();
            ProcessStartInfo sInfo = new ProcessStartInfo();
            process.EnableRaisingEvents = true;
            process.OutputDataReceived += (sender, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    Console.WriteLine(e.Data);
                    //_stdMessages.Enqueue(e.Data);
                }
            };
            process.ErrorDataReceived += (sender, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    Console.Error.WriteLine(e.Data);
                    //_stdErrors.Enqueue(e.Data);
                }
            };
            sInfo.FileName = fileName;
            sInfo.WorkingDirectory = workingDirectory;
            sInfo.Arguments = string.Join(" ", statements);
            sInfo.RedirectStandardError = true;
            sInfo.RedirectStandardOutput = true;
            sInfo.UseShellExecute = false;
            if (noWindow)
            {
                sInfo.WindowStyle = ProcessWindowStyle.Hidden;
            }
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

            do
            {

            } while (!process.HasExited && startTime.Add(timespan) > DateTime.Now);
            int exitCode = !process.HasExited ? 1460 : process.ExitCode; //https://learn.microsoft.com/en-us/windows/win32/debug/system-error-codes--1300-1699-
            if (startTime.Add(timespan) < DateTime.Now)
            {
                Console.Error.WriteLine("Timeout exceeded. Process will now be terminated without completing.");
                if (KillProcess(process, true))
                {
                    Console.WriteLine($"Attempting to kill process {process.ProcessName}, and any child processes.");
                    process.WaitForExit();
                    process.Dispose();
                    process = null;
                }
            }
            if (!validExitCodes.Contains(exitCode))
            {
                throw new ArgumentException(
                    $"Running {fileName} {string.Join(" ", statements)}] was not successful. Exit code was {exitCode}.");
            }
        }

        /// <summary>
        /// Kill the process
        /// </summary>
        /// <param name="process"></param>
        /// <param name="recursiveKill">Will kill all sub-processes of the provided <paramref name="process"/></param>
        /// <returns></returns>
        public static bool KillProcess(System.Diagnostics.Process process, bool recursiveKill = false)
        {
            if (process == null) throw new ArgumentNullException(nameof(process));
            
            if (recursiveKill)
                return KillChildProcess(process.Id, process.MachineName, true);
            

            ManagementScope scope = new ManagementScope($"\\\\{process.MachineName}\\root\\cimv2");
            ObjectQuery query = new ObjectQuery($"SELECT * FROM Win32_Process WHERE ProcessId='{process.Id}'");

            scope.Connect();
            if (!scope.IsConnected)
                throw new ManagementException(
                    $"Failed to connect to the application server. MachineName= {process.MachineName}");
            using (var searcher = new ManagementObjectSearcher(scope, query))
            {
                foreach (var o in searcher.Get())
                {
                    var obj = (ManagementObject)o;
                    obj.InvokeMethod("Terminate", null);
                }
                return true;
            }
        }

        /// <summary>
        /// Kill the process
        /// </summary>
        /// <param name="processId"></param>
        /// <param name="computerName">Computer name where the process is running</param>
        /// <param name="recursiveKill">Will kill all sub-processes of the provided <paramref name="processId"/></param>
        /// <returns></returns>
        public static bool KillChildProcess(int processId, string computerName = null, 
            bool recursiveKill = false)
        {
            
            var machine = string.IsNullOrEmpty(computerName) ? Environment.MachineName : computerName;
            ManagementScope scope = new ManagementScope($"\\\\{machine}\\root\\cimv2");

            scope.Connect();
            if (!scope.IsConnected)
                throw new ManagementException(
                    $"Failed to connect to the application server. MachineName= {machine}");
            ObjectQuery query = new ObjectQuery($"SELECT * FROM Win32_Process WHERE ParentProcessID='{processId}'");
            var childCol = new ManagementObjectSearcher(scope, query).Get();
            if (childCol.Count > 0)
            {
                foreach (var o in childCol)
                {
                    var obj = (ManagementObject)o;
                    UInt32 childId = (UInt32)obj["ProcessId"];
                    if (processId != childId && recursiveKill && !KillChildProcess((int)childId, computerName, true))
                    {
                        throw new SecurityException(
                            $"Unable to stop child process '{obj["ProcessName"]}' with Id '{childId}'");
                    }
                }
            }
            //search for top-level process now, in event that killing child process has forced exit of top-level process
            ObjectQuery curProcess = new ObjectQuery($"SELECT * FROM Win32_Process WHERE ProcessId='{processId}'");
            ManagementObjectCollection curCol = new ManagementObjectSearcher(scope, curProcess).Get();

            foreach (var o in curCol)
            {
                var curObj = (ManagementObject)o;
                curObj.InvokeMethod("Terminate", null);
            }

            return true;
        }

    }
}
