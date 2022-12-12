using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using CsvHelper;
using CommandLineParser.Arguments;
using CsvHelper.Configuration;
using DCOM.Software_Management;
using win_int_info.Serialization;
using WinDevOps;
using WinDevOps.Accounts;
using WinDevOps.WMI;

namespace win_int_info
{
    public class ParsingTarget
    {
        [ValueArgument(typeof(string), 'o', "output", Description = "Define the output folder where files will be created", AllowMultiple = false, ValueOptional = false, Optional = false)]
        public string Output { get; set; }

        [SwitchArgument( 'h', "help", false, Aliases = new string[]{"?"}, Description = "Show help")]
        public bool Help { get; set; }
        
        [ValueArgument(typeof(string), 't', "target", Description = "Name of computer to gather details from", AllowMultiple = false, ValueOptional = true)]
        public string Target { get; set; }

        [ValueArgument(typeof(string), 'u', "user", Description = "Name of user account to use when gathering remote target details.", AllowMultiple = false, ValueOptional = true)]
        public string User { get; set; }

        [ValueArgument(typeof(string), 'p', "password", Description = "Password of the user.", AllowMultiple = false, ValueOptional = true)]
        public string Password { get; set; }

    }

    class Program
    {
        // single-threaded apartment
        //This attribute must be present on the entry point of any application that uses Windows Forms or COM classes such as Shell32;
        //if it is omitted, the Windows components might not work correctly.
        //If the attribute is not present, the application uses the multi-threaded apartment model, which is not supported.
        [STAThread]
        static void Main(string[] args)
        {
            ParsingTarget p = new ParsingTarget();
            CommandLineParser.CommandLineParser parser = new CommandLineParser.CommandLineParser();
            try
            {
                parser.ExtractArgumentAttributes(p);
                parser.ParseCommandLine(args);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
                Console.WriteLine("Press any key to exit....");
                Console.ReadKey();
                return;
            }

            if (p.Help)
            {
                parser.ShowUsageHeader = "Here is how you use the app: ";
                parser.ShowUsageFooter = "Have fun!";
                parser.ShowUsage();
                Console.WriteLine("Press any key to exit....");
                Console.ReadKey();
                return;
            }

            if (string.IsNullOrEmpty(p.Output))
            {
                Console.Error.WriteLine("WARNING: --filepath parameter must be provided.");
                Console.WriteLine("Press any key to exit");
                Console.ReadKey();
                return;
            }

            if (!Directory.Exists(p.Output))
            {
                Console.WriteLine($"Attempting to create output directory: {p.Output}");
                try
                {
                    Directory.CreateDirectory(p.Output);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Exception thrown when creating output directory: {ex.Message}");
                    return;
                }
            }

            NetworkCredential cred = null;
            if (!string.IsNullOrEmpty(p.User))
            {
                cred = new NetworkCredential(p.User, p.Password);
            }
            #region ComputerInfo

            var filePrefix = string.IsNullOrEmpty(p.Target) ? $"{Environment.MachineName}_" : $"{p.Target}_";
            ComputerInformation h = null;
            try
            {
                Console.WriteLine("Retrieving Computer Information settings");
                h = WinDevOps.WMI.WMIUtils.GetComputerInformation(p.Target, cred);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to retrieve Computer Information", ex);
            }

            if(h !=null)
            {
                var serializeResult = SerializeSingleton(h.ComputerSystem, p.Output, $"{filePrefix}computerSystem.csv",
                    typeof(CompPropMap));
                Console.WriteLine("**************************************");
                Console.WriteLine($"Export ComputerSystem: {serializeResult}");
                Console.WriteLine("**************************************");
                serializeResult = SerializeSingleton(h.OperatingSystem, p.Output, $"{filePrefix}operatingSystem.csv",
                    typeof(OsPropMap));
                Console.WriteLine("**************************************");
                Console.WriteLine($"Export operatingSystem: {serializeResult}");
                Console.WriteLine("**************************************");
                serializeResult = SerializeSingleton(h.Chassis, p.Output, $"{filePrefix}chassis.csv", typeof(ChassisPropMap));
                Console.WriteLine("**************************************");
                Console.WriteLine($"Export chassis: {serializeResult}");
                serializeResult = SerializeSingleton(h.Bios, p.Output, $"{filePrefix}bios.csv", typeof(BiosPropMap));
                Console.WriteLine("**************************************");
                Console.WriteLine($"Export bios: {serializeResult}");
                Console.WriteLine("**************************************");
                serializeResult = SerializeArray(h.MemoryProperties, p.Output, $"{filePrefix}memory.csv", typeof(MemoryPropMap));
                Console.WriteLine("**************************************");
                Console.WriteLine($"Export memory: {serializeResult}");
                Console.WriteLine("**************************************");
                serializeResult = SerializeArray(h.MemoryArrayProperties, p.Output, $"{filePrefix}memoryArray.csv",
                    typeof(MemoryArrayPropMap));
                Console.WriteLine("**************************************");
                Console.WriteLine($"Export memoryArray: {serializeResult}");
                Console.WriteLine("**************************************");
                serializeResult = SerializeSingleton(h.ProcessorProperties, p.Output, $"{filePrefix}processor.csv",
                    typeof(ProcPropMap));
                Console.WriteLine("**************************************");
                Console.WriteLine($"Export processor: {serializeResult}");
                serializeResult = SerializeSingleton(h.Time, p.Output, $"{filePrefix}time.csv", null);
                Console.WriteLine("**************************************");
                Console.WriteLine($"Export time: {serializeResult}");
                Console.WriteLine("**************************************");
                serializeResult = SerializeArray(h.NetworkCards, p.Output, $"{filePrefix}networkCards.csv", null);
                Console.WriteLine("**************************************");
                Console.WriteLine($"Export networkCards: {serializeResult}");
                Console.WriteLine("**************************************");
                serializeResult = SerializeArray(h.Disks, p.Output, $"{filePrefix}disks.csv", typeof(DiskPropMap));
                Console.WriteLine("**************************************");
                Console.WriteLine($"Export disks: {serializeResult}");
                Console.WriteLine("**************************************");
            }
            #endregion ComputerInfo

            #region DCOM

            FormatSection("Retrieving Global DCOM settings"); 

            DCOMGlobal gbl = new DCOMGlobal();
            bool r = SerializeSingleton(gbl, p.Output, $"{filePrefix}dcomComputer.csv", typeof(DcomGblMap));
            Console.WriteLine("**************************************");
            Console.WriteLine($"Export DCOM Global Config: {r}");
            Console.WriteLine("**************************************");

            FormatSection("Retrieving all DCOM application settings");
            
            var apps = DCOMApplication.GetDcomApplications(p.Target, null, gbl);
            r = SerializeArray(apps, p.Output, $"{filePrefix}dcomApps.csv", typeof(DcomAppMap));
            Console.WriteLine("**************************************");
            Console.WriteLine($"Export DCOM Application Config: {r}");
            Console.WriteLine("**************************************");
            #endregion DCOM

            #region Security
            FormatSection("Exporting Windows Security");

            var grps = AccountUtils.GetLocalWinUsers(p.Target, cred);
            r = SerializeArray(grps, p.Output, $"{filePrefix}windowsGroups.csv", typeof(WindowsGroupMap));
            Console.WriteLine("**************************************");
            Console.WriteLine($"Export Local Windows Security: {r}");
            Console.WriteLine("**************************************");

            #endregion Security 

            #region Software
            FormatSection("Exporting Software Programs");

            var software = SoftwareUtils.GetSoftware(p.Target);
            r = SerializeArray(software, p.Output, $"{filePrefix}software.csv", typeof(SoftwareProgramMap));
            Console.WriteLine("**************************************");
            Console.WriteLine($"Export Software Programs List: {r}");
            Console.WriteLine("**************************************");

            #endregion

            // The following commands call other cmdline app's using a new process
            #region Windows Firewall
            FormatSection("Exporting Windows Firewall rules");

            var cmd = new string[]
                {"/c", $"netsh advfirewall firewall show rule name=all verbose > \"{Path.Combine(p.Output,$"{filePrefix}firewallRules.txt")}\""};
            try
            {
                WinDevOps.Process.ProcessUtils.InvokeProcess(cmd);
            }
            catch
            {
                Console.WriteLine("Failed to export Firewall Rules");
            }

            #endregion


            #region Windows Services
            FormatSection("Exporting Windows services");

            cmd = new string[]
                {"/c", $"wmic /output:\"{Path.Combine(p.Output,$"{filePrefix}services.csv")}\" service get DisplayName,Name,StartMode,StartName,PathName /format:csv"};
            try
            {
                WinDevOps.Process.ProcessUtils.InvokeProcess(cmd);
            }
            catch
            {
                Console.WriteLine("Failed to export Services");
            }

            #endregion
            
            #region Processes

            FormatSection("Exporting windows processes");

            cmd = new string[]
                {"/c", $"wmic /output:\"{Path.Combine(p.Output,$"{filePrefix}processes.csv")}\" process get /format:csv"};
            try
            {
                WinDevOps.Process.ProcessUtils.InvokeProcess(cmd);
            }
            catch
            {
                Console.WriteLine("Failed to export Processes");
            }

            #endregion

            #region Local Security Policy
            FormatSection("Exporting local security policy");
            cmd = new string[]
                {"/c", $"secedit.exe /export /cfg \"{Path.Combine(p.Output,$"{filePrefix}security-policy.inf")}\""};
            try
            {
                WinDevOps.Process.ProcessUtils.InvokeProcess(cmd);
            }
            catch
            {
                Console.WriteLine("Failed to export Local Security Policy");
            }

            #endregion

            #region MS SQL
            FormatSection("Exporting MS SQL x32");
            var str = string.Format("reg export \"{0}\" \"{1}\" /y ", (object)"HKLM\\SOFTWARE\\Microsoft\\Microsoft SQL Server",
                (object)(Path.Combine(p.Output, $"{filePrefix}Microsoft SQL Server_x32.reg")));
            cmd = new string[]
                {"/c", str};
            try
            {
                WinDevOps.Process.ProcessUtils.InvokeProcess(cmd);
            }
            catch
            {
                Console.WriteLine("Failed to export MSSQL x32 Registry Key");
            }
            FormatSection("Exporting MS SQL x64");
            
            str = string.Format("reg export \"{0}\" \"{1}\" /y ", (object)"HKLM\\SOFTWARE\\Wow6432Node\\Microsoft\\Microsoft SQL Server", 
                (object)(Path.Combine(p.Output, $"{filePrefix}Microsoft SQL Server_x64.reg")));
            cmd = new string[]
                {"/c", str};
            try
            {
                WinDevOps.Process.ProcessUtils.InvokeProcess(cmd);
            }
            catch
            {
                Console.WriteLine("Failed to export MSSQL x64 Registry Key");
            }
            #endregion

            #region ODBC

            FormatSection("Exporting ODBC x32");
            str = string.Format("reg export \"{0}\" \"{1}\" /y ", (object)"HKLM\\SOFTWARE\\ODBC",
                (object)(Path.Combine(p.Output, $"{filePrefix}ODBC_x32.reg")));
            cmd = new string[]
                {"/c", str};
            try
            {
                WinDevOps.Process.ProcessUtils.InvokeProcess(cmd);
            }
            catch
            {
                Console.WriteLine("Failed to export ODBC x32 Registry Key");
            }
            
            FormatSection("Exporting ODBC x64");
            str = string.Format("reg export \"{0}\" \"{1}\" /y ", (object)"HKLM\\SOFTWARE\\WOW6432Node\\ODBC",
                (object)(Path.Combine(p.Output, $"{filePrefix}ODBC_x64.reg")));
            cmd = new string[]
                {"/c", str};
            try
            {
                WinDevOps.Process.ProcessUtils.InvokeProcess(cmd);
            }
            catch
            {
                Console.WriteLine("Failed to export ODBC x64 Registry Key");
            }
            #endregion

            #region AutoLogin
            FormatSection("Exporting AutoLogin registry key");

            str = string.Format("reg export \"{0}\" \"{1}\" /y ", (object)"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon",
                (object)(Path.Combine(p.Output, $"{filePrefix}AutoLogin.reg")));
            cmd = new string[]
                {"/c", str};
            try
            {
                WinDevOps.Process.ProcessUtils.InvokeProcess(cmd);
            }
            catch
            {
                Console.WriteLine("Failed to export AutoLogin Registry Key");
            }
            #endregion AutoLogin

            #region Windows Services
            FormatSection("Exporting network statistics");

            cmd = new string[]
                {"/c", $"netstat -a -o -n > \"{Path.Combine(p.Output,$"{filePrefix}netstat.txt")}\""};
            try
            {
                WinDevOps.Process.ProcessUtils.InvokeProcess(cmd);
            }
            catch
            {
                Console.WriteLine("Failed to network statistics");
            }

            #endregion
            
            #region Interesting Files

            var ora = Environment.GetEnvironmentVariable("TNS_ADMIN");
            if (!string.IsNullOrEmpty(ora))
            {
                FormatSection("Exporting Oracle TNS ORA file");

                try
                {
                    WinDevOps.File_Management.FileUtils.CopyFile(Path.Combine(ora, "tnsnames.ora"), Path.Combine(p.Output, $"{filePrefix}tnsnames.ora"));
                }
                catch
                {
                    Console.Error.WriteLine("Failed to gather Oracle files");
                }
            }

            if (File.Exists($"C:\\ProgramData\\Rockwell\\RSLinx Enterprise\\RSLinxNG.xml"))
            {
                FormatSection("Exporting RSLinx Project file");
                try
                {
                    WinDevOps.File_Management.FileUtils.CopyFile(
                        "C:\\ProgramData\\Rockwell\\RSLinx Enterprise\\RSLinxNG.xml",
                        Path.Combine(p.Output, $"{filePrefix}rslinxNG.xml"));
                }
                catch
                {
                    Console.Error.WriteLine("Failed to gather RSLinx files");
                }
            }

            if (File.Exists($"C:\\ProgramData\\Kepware\\KEPServerEX\\V6\\default.opf"))
            {
                FormatSection("Exporting Kepware Project file");

                WinDevOps.File_Management.FileUtils.CopyFile("C:\\ProgramData\\Kepware\\KEPServerEX\\V6\\default.opf", Path.Combine(p.Output, $"{filePrefix}kepware_default.opf"));
                try
                {
                    WinDevOps.File_Management.FileUtils.CopyFile(
                        "C:\\ProgramData\\Kepware\\KEPServerEX\\V6\\settings.ini",
                        Path.Combine(p.Output, $"{filePrefix}kepware_settings.ini"));
                }
                catch
                {
                    Console.Error.WriteLine("Failed to gather Kepware files");
                }
            }


            #endregion Interesting Files

            #region Create ZipFile

            try
            {
                var root = Directory.GetParent(p.Output) == null ? "C:\\" : Directory.GetParent(p.Output).FullName;
                var outputZip = Path.Combine(root, $"{filePrefix}system_info.zip");
                WinDevOps.File_Management.CompressionUtils.CompressFolder(outputZip, p.Output, true);
            }
            catch
            {
                Console.Error.WriteLine($"Failed to create zip file of exported data. Must be completed manually.");
            }
            #endregion

            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("=========================================");
            Console.WriteLine("=========================================");
            Console.WriteLine("Completed capture and export of settings....");
            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
        }

        /// <summary>
        /// Export object to csv file
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="outputObject">Object to be serialized to csv</param>
        /// <param name="dir">Output directory</param>
        /// <param name="fileName">Output filename. Will be overwritten if exists</param>
        /// <param name="classMapType">Class Map. Leave null for <see cref="CsvClassMap.AutoMap"/></param>
        /// <returns></returns>
        private static bool SerializeSingleton<T>(T outputObject, string dir, string fileName, Type classMapType = null)
        {
            try
            {
                using (var writer = new StreamWriter(Path.Combine(dir, fileName)))
                {
                    var config = new CsvConfiguration();
                    if (classMapType != null) config.RegisterClassMap(classMapType);
                    using (CsvWriter csv = new CsvWriter(writer, config))
                    {
                        csv.WriteHeader<T>();
                        csv.WriteRecord(outputObject);
                        return File.Exists(Path.Combine(dir, fileName));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to serialize {fileName}. {ex.Message}");
                return false;
            }
        }
        /// <summary>
        /// Export object array to csv file
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="outputObject">Object array to be serialized to csv</param>
        /// <param name="dir">Output directory</param>
        /// <param name="fileName">Output filename. Will be overwritten if exists</param>
        /// <param name="classMapType">Class Map. Leave null for <see cref="CsvClassMap.AutoMap"/></param>
        /// <returns></returns>
        private static bool SerializeArray<T>(IEnumerable<T> outputObject, string dir, string fileName, Type classMapType = null)
        {
            try
            {
                using (var writer = new StreamWriter(Path.Combine(dir, fileName)))
                {
                    var config = new CsvConfiguration();
                    if (classMapType != null) config.RegisterClassMap(classMapType);
                    using (CsvWriter csv = new CsvWriter(writer, config))
                    {
                        csv.WriteRecords(outputObject);
                        return File.Exists(Path.Combine(dir, fileName));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed to serialize {fileName}. {ex.Message}");
                return false;
            }

        }

        private static void FormatSection(string message)
        {
            Console.WriteLine("=========================================");
            Console.WriteLine("=========================================");
            Console.WriteLine();
            Console.WriteLine(message);
            Console.WriteLine();
            Console.WriteLine();

        }
    }
}
