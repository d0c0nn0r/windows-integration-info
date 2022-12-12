using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using System.Security.AccessControl;
using System.Security.Principal;
using DCOM.Global;
using Microsoft.Win32;

namespace WinDevOps
{
    /// <summary>
    /// Class that represents DCOM Applications, as they appear in DCOMCNFG
    /// </summary>
    public class DCOMApplication : IEquatable<DCOMApplication>
    {
        private readonly DCOMGlobal _global;

        private List<DCOMAce> _launchPermissions;
        private List<DCOMAce> _accessPermissions;

        /// <summary>
        /// Host where this DCOMApplication setting is installed
        /// </summary>
        public string ComputerName { get; private set; }

        /// <summary>
        /// Application Name
        /// </summary>
        public string AppName { get; internal set; }

        /// <summary>
        /// Application Unique Identifier
        /// </summary>
        public string AppId { get; internal set; }

        /// <summary>
        /// Launch and activation Permissions for this application
        /// </summary>
        public List<DCOMAce> LaunchPermissions
        {
            get
            {
                if (_launchPermissions == null)
                {
                    return _global?.DefaultLaunchAndActivation;
                }

                return _launchPermissions;
            }
            internal set => _launchPermissions = value;
        }

        /// <summary>
        /// If true, Application uses global default settings, defined on <see cref="DCOMGlobal"/>
        /// If false, this is customized for this application
        /// </summary>
        public bool LaunchPermissionsUseDefault => _launchPermissions == null;

        /// <summary>
        /// Access Permissions for this application
        /// </summary>
        public List<DCOMAce> AccessPermissions
        {
            get
            {
                if (_accessPermissions == null)
                {
                    return _global?.DefaultAccessPermissions;
                }

                return _accessPermissions;
            }
            internal set => _accessPermissions = value;
        }

        /// <summary>
        /// If true, Application uses global default settings, defined on <see cref="DCOMGlobal"/>
        /// If false, this is customized for this application
        /// </summary>
        public bool AccessPermissionsUseDefault => _accessPermissions == null;

        /// <summary>
        /// Authentication level required by this DCOM app
        /// </summary>
        public RPC_C_AUTHN AuthenticationLevel { get; internal set; }

        /// <summary>
        /// Service name, as would appear in services applet. Empty is application does not run as a service
        /// </summary>
        public string ServiceName { get; internal set; }

        /// <summary>
        /// Service startup type, as would appear in services applet. Empty is application does not run as a service
        /// </summary>
        public ServiceStart ServiceStartup { get; internal set; }

        /// <summary>
        /// Account that windows service, or DCOM Application is configured to run under.
        /// Value is Empty if application does not run as a service, and instead run as a process
        /// using the Interactive or Launching User
        /// </summary>
        public string RunAs { get; internal set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        internal DCOMApplication(DCOMGlobal global = null)
        {
            ComputerName = Environment.MachineName;
            
            _global = global ?? new DCOMGlobal();
        }

        internal DCOMApplication(string computerName = null, 
            DCOMGlobal global = null)
        {
            ComputerName = String.IsNullOrEmpty(computerName) ? Environment.MachineName : computerName;
            _global = global ?? new DCOMGlobal();
        }


        /// <summary>
        /// Apply changes to <seealso cref="AccessPermissions"/> Access Control list.
        /// Change in privileges are applied to only the <paramref name="account"></paramref>.
        /// If account is not in Access Control List, then it is added.
        /// </summary>
        /// <param name="account">The account to apply the changes to</param>
        /// <param name="rights">The privileges to assign to the account</param>
        /// <param name="permissionType">Privileges can be either ALLOWED or DENIED</param>
        /// <returns></returns>
        public bool SetLaunchRights(SecurityIdentifier account, MachineDComAccessRights[] rights,
            AccessControlType permissionType)
        {
            var res = DcomUtils.AddApplicationRights(DcomSecurityTypes.Launch, account, rights, permissionType,
                AppId,
                ComputerName);
            var app = GetDcomApplications(ComputerName, AppName)[0];
            LaunchPermissions = app?.LaunchPermissions;

            return res;
        }
        

        /// <summary>
        /// Remove <paramref name="account"/> from the <seealso cref="LaunchPermissions"/> Access Control List
        /// </summary>
        /// <param name="account">The account to remove from the ACE</param>
        /// <returns></returns>
        public bool RemoveLaunchRights(SecurityIdentifier account)
        {
            var res = DcomUtils.RemoveApplicationRights(DcomSecurityTypes.Launch, account, AppId,
                ComputerName);
            var app = GetDcomApplications(ComputerName, AppName)[0];
            LaunchPermissions = app?.LaunchPermissions;

            return res;
        }
        

        /// <summary>
        /// Apply changes to <seealso cref="LaunchPermissions"/> Access Control list.
        /// Change in privileges are applied to only the <paramref name="account"></paramref>.
        /// If account is not in Access Control List, then it is added.
        /// </summary>
        /// <param name="account">The account to apply the changes to</param>
        /// <param name="rights">The privileges to assign to the account</param>
        /// <param name="permissionType">Privileges can be either ALLOWED or DENIED</param>
        /// <returns></returns>
        public bool SetAccessRights(SecurityIdentifier account, MachineDComAccessRights[] rights,
            AccessControlType permissionType)
        {
            var res = DcomUtils.AddApplicationRights(DcomSecurityTypes.Access, account, rights, permissionType,
                AppId,
                ComputerName);
            var app = GetDcomApplications(ComputerName, AppName)[0];
            AccessPermissions = app?.AccessPermissions;

            return res;
        }
        
        
        /// <summary>
        /// Export DCOM application registry key as a .reg file, that can be used for import/export
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        /// <remarks>WARNING. This function will rename any files that pre-exist at the given <paramref name="filePath"></paramref></remarks>
        private bool CreateRegistryKeyBackup(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException(nameof(filePath));
            }
            
            using (var id = WindowsIdentity.GetCurrent())
            {
                var p = new WindowsPrincipal(id);
                if (!p.IsInRole(WindowsBuiltInRole.Administrator))
                {
                    throw new SecurityException("User must be running as an administrator");
                }
            }

            if (!File.Exists(filePath))
                return RegistryFunctions.ExportRegistryKey(
                    $"HKEY_CLASS_ROOT\\AppID\\{AppId}", filePath);

            var fName = FileOperations.IncrementFileName(filePath);
            File.Move(filePath, fName);

            return RegistryFunctions.ExportRegistryKey($"HKEY_CLASS_ROOT\\AppID\\{AppId}", filePath);
        }

        /// <summary>
        /// Import DCOM application registry key backup from a .reg file, that was created using <see cref="CreateRegistryKeyBackup"/>
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        /// <remarks>WARNING. This function will overwrite local REGISTRY key. This is a </remarks>
        public static bool ImportRegistryKeyBackup(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException(nameof(filePath));
            }

            using (var id = WindowsIdentity.GetCurrent())
            {
                var p = new WindowsPrincipal(id);
                if (!p.IsInRole(WindowsBuiltInRole.Administrator))
                {
                    throw new SecurityException("User must be running as an administrator");
                }
            }

            if (!File.Exists(filePath))
                throw new FileNotFoundException($"File not found", filePath);

            return RegistryFunctions.ImportRegistryKey(filePath);
        }

        /// <summary>
        /// Update the <see cref="AccessPermissions"/> of this application, by copying the ACL from <paramref name="copyFrom"/>
        /// defined on the other application.
        /// To replace ACL completely, replacing/removing any of the current settings, set <paramref name="overwrite"/> to true.
        /// </summary>
        /// <param name="copyFrom">The <see cref="DCOMApplication"/> to copy <see cref="AccessPermissions"/> from</param>
        /// <param name="overwrite">Set to true, to replace and remove privileges defined locally, entirely with those in <paramref name="copyFrom"/>.
        /// If false, missing accounts are added, and privileges for matching accounts are updated to match <paramref name="copyFrom"/></param>
        /// <param name="type">The Access Control List to update. Can be either Launch or Access. .</param>
        /// <exception cref="NotSupportedException">Thrown if <paramref name="type"/> value provided is Config. This is not supported for ACE copying</exception>
        /// <returns></returns>
        /// <exception cref="AggregateException">Exceptions representing ACE rules that failed to be copied successfully.</exception>
        /// <remarks>
        /// If the <paramref name="copyFrom"/> application Access or Launch Permissions are set to "use default", then
        /// this function will call <see cref="UseDefaultAccessPermissions"/> and <see cref="UseDefaultLaunchPermissions"/> methods.
        /// This does not guarantee that DCOM Application securities between different computers match.
        /// Using Default means security permissions are inherited from the corresponding Default Access Control List defined on <see cref="DCOMGlobal"/>.
        /// </remarks>
        public bool CopyAccessControlList(DCOMApplication copyFrom, bool overwrite = false,
            DcomSecurityTypes type = DcomSecurityTypes.Access)
        {
            if (type == DcomSecurityTypes.Config)
            {
                throw new NotSupportedException(
                    "DcomSecurityTypes.Config is not supported for copying Access Control Lists");
            }

            var retList = new List<bool>();
            if (copyFrom.AccessPermissionsUseDefault || copyFrom.LaunchPermissionsUseDefault)
            {
                if (copyFrom.AccessPermissionsUseDefault)
                {
                    try
                    {
                        var r1 = UseDefaultAccessPermissions();
                        if (!r1)
                            throw new Exception(
                                $"{ComputerName} : Failed to reset Application '{AppName}' Access Permissions to default");
                        retList.Add(r1);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Failed to reset AccessPermissions to Use Default", ex);
                    }
                }

                if (copyFrom.LaunchPermissionsUseDefault)
                {
                    try
                    {
                        var r1 = UseDefaultLaunchPermissions();
                        if (!r1)
                            throw new Exception(
                                $"{ComputerName} : Failed to reset Application '{AppName}' Launch Permissions to default");
                        retList.Add(r1);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Failed to reset LaunchPermissions to Use Default", ex);
                    }
                }
            }
            else
            {
                var toRemove = type == DcomSecurityTypes.Access
                    ? GetMismatchedAccessPermissions(copyFrom)
                    : GetMismatchedLaunchPermissions(copyFrom);
                var toAdd = type == DcomSecurityTypes.Access
                    ? copyFrom.GetMismatchedAccessPermissions(this)
                    : copyFrom.GetMismatchedLaunchPermissions(this);
                foreach (var dcomAce in toAdd)
                {
                    var r = DcomUtils.AddApplicationRights(dcomAce.SecurityType, dcomAce.SID,
                        dcomAce.ConvertToMachineRightsList().ToArray(), dcomAce.AccessType, AppId, ComputerName);
                    if (!r)
                        throw new Exception(
                            $"Failed to apply DCOMAce for user '{dcomAce.User}', Type: {dcomAce.AccessType}, Category: {dcomAce.Category}, SecurityType: {dcomAce.SecurityType}");
                    retList.Add(r);
                }

                if (overwrite)
                {
                    foreach (var dcomAce in toRemove)
                    {
                        var r = DcomUtils.RemoveApplicationRights(dcomAce.SecurityType, dcomAce.SID, AppId,
                            ComputerName);
                        if (!r)
                            throw new Exception(
                                $"Failed to remove DCOMAce for user '{dcomAce.User}', Type: {dcomAce.AccessType}, Category: {dcomAce.Category}, SecurityType: {dcomAce.SecurityType}");
                        retList.Add(r);
                    }
                }
            }

            var app = GetDcomApplication(new List<string>(1) { AppId }, ComputerName, null)[0];
            AccessPermissions = app?.AccessPermissions;
            
            return retList.TrueForAll(a => a);
        }

        /// <summary>
        /// Remove <paramref name="account"/> from the <seealso cref="AccessPermissions"/> Access Control List
        /// </summary>
        /// <param name="account">The account to remove from the ACE</param>
        /// <returns></returns>
        public bool RemoveAccessRights(SecurityIdentifier account)
        {
            var res = DcomUtils.RemoveApplicationRights(DcomSecurityTypes.Access, account, AppId,
                ComputerName);
            var app = GetDcomApplication(new List<string>(1) { AppId }, ComputerName, null)[0];
            AccessPermissions = app?.AccessPermissions;

            return res;
        }
        
        /// <summary>
        /// Set this application to Use Default Launch Permissions
        /// </summary>
        /// <returns></returns>
        public bool UseDefaultLaunchPermissions()
        {
            var res = DcomUtils.ResetApplicationRightsToDefault(DcomSecurityTypes.Launch, AppId,
                ComputerName);
            if (res)
            {
                var app = GetDcomApplication(new List<string>(1) { AppId }, ComputerName, null)[0];
                LaunchPermissions = app?.LaunchPermissions;
            }

            return res;
        }

        /// <summary>
        /// Set this application to Use Default Access Permissions
        /// </summary>
        /// <returns></returns>
        public bool UseDefaultAccessPermissions()
        {
            var res = DcomUtils.ResetApplicationRightsToDefault(DcomSecurityTypes.Access, AppId,
                ComputerName);
            if (res)
            {
                var app = GetDcomApplication(new List<string>(1) { AppId }, ComputerName,  null)[0];
                AccessPermissions = app?.AccessPermissions;
            }

            return res;
        }

        /// <summary>
        /// Retrieve DCOM applications from the target computer registry
        /// </summary>
        /// <param name="computerName">Name of the computer to retrieve applications from</param>
        /// <param name="appName">Filter returned applications. Does not support wildcards.</param>
        /// <param name="computerDcom"><see cref="DCOMGlobal"/> object that has all global defaults that govern the DcomApplications</param>
        /// <returns></returns>
        public static List<DCOMApplication> GetDcomApplications(string computerName = null,
            string appName = null, DCOMGlobal computerDcom = null)
        {
            var apps = new List<DCOMApplication>();
            
            RegistryKey rootKey = string.IsNullOrEmpty(computerName)
                ? Registry.ClassesRoot
                : RegistryKey.OpenRemoteBaseKey(RegistryHive.ClassesRoot, computerName);
            
            using (RegistryKey registryKey = rootKey.OpenSubKey("AppID"))
            {
                if (registryKey == null)
                    throw new InvalidOperationException(
                        $"{computerName} : The registry key 'HKEY_CLASS_ROOT\\AppID' was not found.");
                foreach(var subKeyName in registryKey.GetSubKeyNames())
                    //Parallel.ForEach(registryKey.GetSubKeyNames(), options, ((subKeyName, state) =>
                {
                    DCOMApplication app = new DCOMApplication(computerName, computerDcom);
                    Console.WriteLine(
                        $"{computerName} : Attempting to process application {subKeyName}");

                    using (RegistryKey subKey = registryKey.OpenSubKey(subKeyName))
                    {
                        string sdefault = (string) subKey.GetValue(null);
                        try
                        {
                            byte[] launchPermis = (byte[]) subKey.GetValue("LaunchPermission");
                            byte[] accessPermis = (byte[]) subKey.GetValue("AccessPermission");
                            int? authenticationLevel = subKey.GetValue("AuthenticationLevel") as int?;
                            string sService = (string) subKey.GetValue("LocalService");
                            if (String.IsNullOrEmpty(sdefault))
                            {
                                Console.WriteLine(
                                    $"{computerName}\\{subKeyName} : Necessary Registry keys did not exist. Ignoring application");
                                continue; //default doesn't exist, go to next
                            }

                            //If appName=null, process regKey. Else, if current App matches any of the names in the appNames list (including wildcards)
                            if (appName == null || 
                                string.Equals(appName, sdefault, StringComparison.OrdinalIgnoreCase ))
                            {
                                Console.WriteLine(
                                    $"{computerName}\\{sdefault} : Starting capture of DCOM settings");
                                //add to return list
                                app.AppName = sdefault;
                                app.AppId = subKeyName;
                                app.ServiceName = sService;

                                if (sService != null)
                                {
                                    //Get Service info from Registry

                                    using (RegistryKey serviceKey = string.IsNullOrEmpty(computerName)
                                        ? Registry.LocalMachine.OpenSubKey("SYSTEM")
                                            ?.OpenSubKey("CurrentControlSet")
                                            ?.OpenSubKey("Services")
                                            ?.OpenSubKey(sService)
                                        : RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, computerName)
                                            .OpenSubKey("SYSTEM")
                                            ?.OpenSubKey("CurrentControlSet")
                                            ?.OpenSubKey("Services")
                                            ?.OpenSubKey(sService))
                                    {
                                        if (serviceKey != null)
                                        {
                                            string logOnAs = (string) serviceKey.GetValue("ObjectName");
                                            int? srvStart = serviceKey.GetValue("Start") as int?;
                                            if (srvStart != null)
                                            {
                                                app.ServiceStartup = (ServiceStart) srvStart;
                                            }

                                            app.RunAs = logOnAs;
                                        }
                                    }
                                }
                                else
                                {
                                    var runAs = (string) subKey.GetValue("RunAs");
                                    //If no Service Key exists, then App is either running as Interactive User, or as a specific User
                                    //if null/not exists, then is "Launching User"
                                    app.RunAs = runAs ?? "Launching User" ;
                                }

                                if (launchPermis != null)
                                {
                                    RawSecurityDescriptor newDescriptor =
                                        new RawSecurityDescriptor(launchPermis, 0);
                                    app.LaunchPermissions = (DcomUtils.DeserializeACLtoACE(
                                        DcomSecurityTypes.Launch,
                                        newDescriptor, computerName));
                                }

                                if (accessPermis != null)
                                {
                                    RawSecurityDescriptor newDescriptor =
                                        new RawSecurityDescriptor(accessPermis, 0);
                                    app.AccessPermissions = (DcomUtils.DeserializeACLtoACE(
                                        DcomSecurityTypes.Access,
                                        newDescriptor, computerName));
                                }

                                if (authenticationLevel != null)
                                {
                                    RPC_C_AUTHN doc = (RPC_C_AUTHN) authenticationLevel;
                                    app.AuthenticationLevel = (RPC_C_AUTHN) authenticationLevel;
                                }
                                else
                                {
                                    app.AuthenticationLevel = RPC_C_AUTHN.Default;
                                }

                                Console.WriteLine(
                                    $"{computerName}\\{sdefault} : Completed capture of DCOM settings");
                                apps.Add(app);
                            }
                        }
                        catch (Exception ex)
                        {
                            var s = string.IsNullOrEmpty(sdefault) ? subKeyName : sdefault;
                            Console.Error.WriteLine($"{computerName} : Skipping application {s}. Error while processing:\r\n{ex.Message}");
                            //traditionally would throw EXCEPTION, but in this case we'll just keep going
                            //throw new Exception(
                            //    $"{computerName} : Error while processing application {s}.\r\n{ex.Message}", ex);
                        }
                    }
                }
            }

            return apps;
        
        }

        /// <summary>
        /// Retrieve DCOM applications from the target computer registry
        /// </summary>
        /// <param name="appIds">Exact AppID's of the DCOM Application to read from the registry</param>
        /// <param name="computerName">Name of the computer to retrieve applications from</param>
        /// <param name="computerDcom"><see cref="DCOMGlobal"/> object that has all global defaults that govern the DcomApplications</param>
        /// <returns></returns>
        public static List<DCOMApplication> GetDcomApplication(IList<string> appIds, string computerName = null,
            DCOMGlobal computerDcom = null)
        {
            if (appIds == null || appIds.Count < 1) throw new ArgumentNullException(nameof(appIds));

            var verboseMessages = new List<string>();
            var exceptions = new List<Exception>();
            var apps = new List<DCOMApplication>();
            
            RegistryKey rootKey = string.IsNullOrEmpty(computerName)
                ? Registry.ClassesRoot
                : RegistryKey.OpenRemoteBaseKey(RegistryHive.ClassesRoot, computerName);
            
            using (RegistryKey registryKey = rootKey.OpenSubKey("AppID"))
            {
                if (registryKey == null)
                    throw new InvalidOperationException(
                        $"{computerName} : The registry key 'HKEY_CLASS_ROOT\\AppID' was not found.");
                var keys = new List<string>();
                foreach (var k in registryKey.GetSubKeyNames())
                {
                    for (int i = 0; i < appIds.Count; i++)
                    {
                        if(appIds[i].Equals(k, StringComparison.OrdinalIgnoreCase))
                        {
                            keys.Add(k);
                            break;
                        }
                    }
                }
                //var keys = registryKey.GetSubKeyNames()
                //    .Where(n => appIds.Contains(n, StringComparer.OrdinalIgnoreCase));
                foreach(var subKeyName in keys)
                {
                    
                    DCOMApplication app = new DCOMApplication(computerName, computerDcom);
                    verboseMessages.Add(
                        $"{computerName} : Attempting to process application {subKeyName}");

                    using (RegistryKey subKey = registryKey.OpenSubKey(subKeyName))
                    {
                        string sdefault = (string)subKey.GetValue(null);
                        try
                        {
                            byte[] launchPermis = (byte[])subKey.GetValue("LaunchPermission");
                            byte[] accessPermis = (byte[])subKey.GetValue("AccessPermission");
                            int? authenticationLevel = subKey.GetValue("AuthenticationLevel") as int?;
                            string sService = (string)subKey.GetValue("LocalService");
                            if (String.IsNullOrEmpty(sdefault))
                            {
                                verboseMessages.Add(
                                    $"{computerName}\\{subKeyName} : Necessary Registry keys did not exist. Ignoring application");
                                continue; //default doesn't exist, go to next
                            }

                            verboseMessages.Add(
                                $"{computerName}\\{sdefault} : Starting capture of DCOM settings");
                            //add to return list
                            app.AppName = sdefault;
                            app.AppId = subKeyName;
                            app.ServiceName = sService;

                            if (sService != null)
                            {
                                //Get Service info from Registry

                                using (RegistryKey serviceKey = string.IsNullOrEmpty(computerName)
                                    ? Registry.LocalMachine.OpenSubKey("SYSTEM")
                                        ?.OpenSubKey("CurrentControlSet")
                                        ?.OpenSubKey("Services")
                                        ?.OpenSubKey(sService)
                                    : RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, computerName)
                                        .OpenSubKey("SYSTEM")
                                        ?.OpenSubKey("CurrentControlSet")
                                        ?.OpenSubKey("Services")
                                        ?.OpenSubKey(sService))
                                {
                                    if (serviceKey != null)
                                    {
                                        string logOnAs = (string)serviceKey.GetValue("ObjectName");
                                        int? srvStart = serviceKey.GetValue("Start") as int?;
                                        if (srvStart != null)
                                        {
                                            app.ServiceStartup = (ServiceStart)srvStart;
                                        }

                                        app.RunAs = logOnAs;
                                    }
                                }
                            }
                            else
                            {
                                var runAs = (string)subKey.GetValue("RunAs");
                                //If no Service Key exists, then App is either running as Interactive User, or as a specific User
                                //if null/not exists, then is "Launching User"
                                app.RunAs = runAs ?? "Launching User";
                            }

                            if (launchPermis != null)
                            {
                                RawSecurityDescriptor newDescriptor =
                                    new RawSecurityDescriptor(launchPermis, 0);
                                app.LaunchPermissions = (DcomUtils.DeserializeACLtoACE(
                                    DcomSecurityTypes.Launch,
                                    newDescriptor, computerName));
                            }

                            if (accessPermis != null)
                            {
                                RawSecurityDescriptor newDescriptor =
                                    new RawSecurityDescriptor(accessPermis, 0);
                                app.AccessPermissions = (DcomUtils.DeserializeACLtoACE(
                                    DcomSecurityTypes.Access,
                                    newDescriptor, computerName));
                            }

                            if (authenticationLevel != null)
                            {
                                RPC_C_AUTHN doc = (RPC_C_AUTHN)authenticationLevel;
                                app.AuthenticationLevel = (RPC_C_AUTHN)authenticationLevel;
                            }
                            else
                            {
                                app.AuthenticationLevel = RPC_C_AUTHN.Default;
                            }

                            verboseMessages.Add(
                                $"{computerName}\\{sdefault} : Completed capture of DCOM settings");
                            apps.Add(app);
                        }
                        catch (Exception ex)
                        {
                            var s = string.IsNullOrEmpty(sdefault) ? subKeyName : sdefault;
                            Console.Error.WriteLine($"{computerName} : Skipping application {s}. Error while processing:\r\n{ex.Message}");
                            //exceptions.Add(new Exception(
                            //    $"{computerName} : Error while processing application {s}.\r\n{ex.Message}", ex));
                        }

                    }
                }
            }

            return apps;
        
        }

        /// <inheritdoc />
        public bool Equals(DCOMApplication other)
        {
            StringEqualNullEmptyComparer comparer = new StringEqualNullEmptyComparer();
            return comparer.Equals(AppId, other.AppId, StringComparison.OrdinalIgnoreCase) &&
                   comparer.Equals(AppName, other.AppName, StringComparison.OrdinalIgnoreCase) &&
                   comparer.Equals(RunAs, other.RunAs, StringComparison.OrdinalIgnoreCase) &&
                   comparer.Equals(ServiceName, other.ServiceName, StringComparison.OrdinalIgnoreCase) &&
                   ServiceStartup.Equals(other.ServiceStartup) &&
                   AuthenticationLevel.Equals(other.AuthenticationLevel) &&
                   LaunchPermissionEquality(other) && AccessPermissionEquality(other);
        }

        /// <summary>
        /// Compare <seealso cref="LaunchPermissions"/> between two different DCOM apps, and see it they're the same.
        /// Equality is limited to
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public bool LaunchPermissionEquality(DCOMApplication other)
        {
            if (LaunchPermissions == null || LaunchPermissions.Count <= 0)
            {
                if (other.LaunchPermissions == null || other.LaunchPermissions.Count <= 0)
                {
                    return true;
                }

                return false;
            }

            var cnt = new Dictionary<DCOMAce, int>();
            foreach (var s in LaunchPermissions)
            {
                if (cnt.ContainsKey(s))
                {
                    cnt[s]++; //make dictionary value '1'. will be reduced to 0 by logic of looping 2nd list
                }
                else
                {
                    cnt.Add(s, 1);
                }
            }

            foreach (var s in other.LaunchPermissions)
            {
                if (cnt.ContainsKey(s))
                {
                    if (cnt[s].Equals(s))
                    {
                        cnt[s]--;
                    }
                }
                else
                {
                    return false; //found mismatch. Short-circuit out 'false' result
                }
            }

            var lst = new List<int>(cnt.Values);
            return lst.TrueForAll(c => c == 0);
            //return cnt.Values.All(c => c == 0); //all processed. See if any from source list that did not get matched
        }

        /// <summary>
        /// Get unique list of the <seealso cref="LaunchPermissions"/> objects of that don't match the LaunchPermissions of the <paramref name="other"/> object.
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        /// <remarks>
        /// Return list does not contains any <see cref="DCOMAce"/> objects from the <paramref name="other"/> object
        /// </remarks>
        public List<DCOMAce> GetMismatchedLaunchPermissions(DCOMApplication other)
        {
            return GetMismatched(other, DcomSecurityTypes.Launch);
        }

        /// <summary>
        /// Get unique list of the <seealso cref="AccessPermissions"/> objects of that don't match the LaunchPermissions of the <paramref name="other"/> object.
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        /// <remarks>
        /// Return list does not contains any <see cref="DCOMAce"/> objects from the <paramref name="other"/> object
        /// </remarks>
        public List<DCOMAce> GetMismatchedAccessPermissions(DCOMApplication other)
        {
            return GetMismatched(other, DcomSecurityTypes.Access);
        }

        /// <summary>
        /// Get Unique list of <see cref="DCOMAce"/> rules by comparing the Launch or Access permissions between objects
        /// </summary>
        /// <param name="other"></param>
        /// <param name="type"></param>
        /// <exception cref="NotSupportedException">Thrown if <paramref name="type"/> is set to CONFIG</exception>
        /// <returns></returns>
        /// <remarks>
        /// Return list does not contains any <see cref="DCOMAce"/> objects from the <paramref name="other"/> object
        /// </remarks>
        private List<DCOMAce> GetMismatched(DCOMApplication other, DcomSecurityTypes type)
        {
            if (type == DcomSecurityTypes.Config)
            {
                throw new NotSupportedException(nameof(type));
            }

            List<DCOMAce> rules = new List<DCOMAce>();
            switch (type)
            {
                case DcomSecurityTypes.Access:
                    if (AccessPermissions == null || AccessPermissions.Count <= 0) return new List<DCOMAce>();
                    rules = AccessPermissions;
                    break;
                case DcomSecurityTypes.Launch:
                    if (LaunchPermissions == null || LaunchPermissions.Count <= 0) return new List<DCOMAce>();
                    rules = LaunchPermissions;
                    break;
            }

            if (AccessPermissions == null || AccessPermissions.Count <= 0) return new List<DCOMAce>();
            var cnt = new Dictionary<DCOMAce, int>();
            foreach (var s in rules)
            {
                if (cnt.ContainsKey(s))
                {
                    cnt[s]++; //make dictionary value '1'. will be reduced to 0 by logic of looping 2nd list
                }
                else
                {
                    cnt.Add(s, 1);
                }
            }

            foreach (var s in other.AccessPermissions)
            {
                if (cnt.ContainsKey(s))
                {
                    if (cnt[s].Equals(s))
                    {
                        cnt[s]--;
                    }
                }
            }

            var lst = new List<DCOMAce>();
            foreach (var key in cnt)
            {
                if(key.Value == 1) lst.Add(key.Key);
            }

            return lst;
            //return cnt.Where(c => c.Value == 1).Select(c => c.Key)
            //    .ToList(); //all processed. See if any from source list that did not get matched
        }

        /// <summary>
        /// Compare <seealso cref="AccessPermissions"/> between two different DCOM apps, and see it they're the same
        /// </summary>s
        /// <param name="other"></param>
        /// <returns></returns>
        public bool AccessPermissionEquality(DCOMApplication other)
        {
            if (AccessPermissions == null || AccessPermissions.Count <= 0)
            {
                if (other.AccessPermissions == null || other.AccessPermissions.Count <= 0)
                {
                    return true;
                }

                return false;
            }

            var cnt = new Dictionary<DCOMAce, int>();
            foreach (var s in AccessPermissions)
            {
                if (cnt.ContainsKey(s))
                {
                    cnt[s]++; //make dictionary value '1'. will be reduced to 0 by logic of looping 2nd list
                }
                else
                {
                    cnt.Add(s, 1);
                }
            }

            foreach (var s in other.AccessPermissions)
            {
                if (cnt.ContainsKey(s))
                {
                    if (cnt[s].Equals(s))
                    {
                        cnt[s]--;
                    }
                }
                else
                {
                    return false; //found mismatch. Short-circuit out 'false' result
                }
            }

            var lst = new List<int>(cnt.Values);
            return lst.TrueForAll(c => c == 0);
            //return cnt.Values.All(c => c == 0); //all processed. See if any from source list that did not get matched
        }

        /// <inheritdoc />
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;

            if (obj.GetType() != GetType()) return false;
            DCOMApplication other = (DCOMApplication) obj;
            return Equals(other);
        }

        /// <inheritdoc />
        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 13;
                hash = (hash * 7) ^ AppName.GetHashCode();
                hash = (hash * 7) ^ AppId.GetHashCode();
                return hash;
            }
        }

        /// <summary>
        /// Set Authentication Level in the registry for the DCOMApp
        /// </summary>
        /// <param name="authentication"></param>
        /// <returns></returns>
        private bool UpdateAuthentication(RPC_C_AUTHN authentication)
        {
            RegistryKey rootKey = string.IsNullOrEmpty(ComputerName)
                ? Registry.ClassesRoot
                : RegistryKey.OpenRemoteBaseKey(RegistryHive.ClassesRoot, ComputerName);
            using (RegistryKey registryKey = rootKey.OpenSubKey("AppID"))
            {
                if (registryKey == null)
                    throw new InvalidOperationException(
                        $"{ComputerName} : The registry key 'HKEY_CLASS_ROOT\\AppID' was not found.");

                using (RegistryKey subKey = registryKey.OpenSubKey(AppId))
                {
                    if (subKey == null)
                        throw new ArgumentException(
                            $"{ComputerName} : The registry key 'HKEY_CLASS_ROOT\\AppID\\{AppId}' was not found.");

                    if (authentication == RPC_C_AUTHN.Default)
                    {
                        //remove key to apply default
                        subKey.DeleteValue("AuthenticationLevel", false);
                    }
                    else
                    {
                        //key not found
                        subKey.SetValue("AuthenticationLevel", (int)authentication, RegistryValueKind.DWord);
                    }

                    return true;
                }
            }
        
        }

        /// <summary>
        /// Set RunAs Level in the registry for the DCOMApp
        /// </summary>
        /// <param name="runAs">Change DCOM application runas property.
        /// Only supported for DCOM applications that are NOT windows services.
        /// Therefore, only the following values are allowed; "interactive user", "launching user"
        /// </param>
        /// <returns></returns>
        private bool UpdateRunAs(DcomApplicationRunAs runAs)
        {
            string applyValue = string.Empty;
            switch (runAs)
            {
                case DcomApplicationRunAs.InteractiveUser:
                    applyValue = "Interactive User";
                    break;
                //if launching user, then no registry key will exist
            }
            
            RegistryKey rootKey = string.IsNullOrEmpty(ComputerName)
                ? Registry.ClassesRoot
                : RegistryKey.OpenRemoteBaseKey(RegistryHive.ClassesRoot, ComputerName);
            using (RegistryKey registryKey = rootKey.OpenSubKey("AppID"))
            {
                if (registryKey == null)
                    throw new InvalidOperationException(
                        $"{ComputerName} : The registry key 'HKEY_CLASS_ROOT\\AppID' was not found.");

                using (RegistryKey subKey = registryKey.OpenSubKey(AppId))
                {
                    if (subKey == null)
                        throw new ArgumentException(
                            $"{ComputerName} : The registry key 'HKEY_CLASS_ROOT\\AppID\\{AppId}' was not found.");

                    if (runAs == DcomApplicationRunAs.LaunchingUser)
                    {
                        //remove key to apply default
                        subKey.DeleteValue("RunAs", false);
                    }
                    else if(runAs == DcomApplicationRunAs.InteractiveUser)
                    {
                        //key not found
                        subKey.SetValue("RunAs", applyValue, RegistryValueKind.String);
                    }

                    return true;
                }
            }
        
        }

        /// <summary>
        /// Convert a string value to the corresponding <see cref="DcomApplicationRunAs"/> enum
        /// </summary>
        /// <param name="runAs"></param>
        /// <returns></returns>
        private DcomApplicationRunAs ResolveRunAs(string runAs)
        {
            switch (runAs.ToLower())
            {
                case "interactive user":
                    return DcomApplicationRunAs.InteractiveUser;
                case "launching user":
                    return DcomApplicationRunAs.LaunchingUser;
                default:
                    throw new ArgumentException(
                        $"Argument cannot be converted to a {nameof(DcomApplicationRunAs)} enum type");
                //if launching user, then no registry key will exist
            }
        }
        
    }
}