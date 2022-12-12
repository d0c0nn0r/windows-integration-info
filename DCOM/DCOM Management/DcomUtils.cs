using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using Microsoft.Win32;

namespace WinDevOps
{
    /// <summary>
    /// Manage Dcom Permissions on a computer.
    /// </summary>
    public static class DcomUtils
    {
        /// <summary>
        /// Given <paramref name="securityDescriptor"/>, convert this to DCOM security Access Control List
        /// </summary>
        /// <param name="type"></param>
        /// <param name="securityDescriptor"></param>
        /// <param name="computerName"></param>
        /// <param name="category"></param>
        /// <returns></returns>
        public static List<DCOMAce> DeserializeACLtoACE(DcomSecurityTypes type,
            RawSecurityDescriptor securityDescriptor, string computerName = null, DcomPermissionOption category = DcomPermissionOption.None)
        {
            var dcomAces = new List<DCOMAce>();
            foreach (var genericAce in securityDescriptor.DiscretionaryAcl)
            {
                var dl = (CommonAce)genericAce;
                string user;
                AccessControlType accessControlType = (AccessControlType)dl.AceType;
                List<string> access = new List<string>();
                var dcomAce = new DCOMAce(type, category, dl.SecurityIdentifier);
                try
                {
                    user = dl.SecurityIdentifier
                        .Translate(typeof(NTAccount)).Value;  //convert SID to Readable format
                }
                catch
                {
                    try
                    {
                        uint accountSize = 1024;
                        uint domainSize = 1024;

                        var account = new StringBuilder(1024, 1024);
                        var domain = new StringBuilder(1024, 1024);
                        byte[] bytes = new byte[dl.SecurityIdentifier.BinaryLength];
                        dl.SecurityIdentifier.GetBinaryForm(bytes, 0);
                        var res = string.IsNullOrEmpty(computerName)
                            ? LookupAccountSid(null, bytes,
                                account, ref accountSize, domain, ref domainSize, out _)
                            : LookupAccountSid(computerName,
                                bytes,
                                account, ref accountSize, domain, ref domainSize, out _);
                        user = res ? $"{domain}\\{account}" : dl.SecurityIdentifier.Value;
                    }
                    catch
                    {
                        user = dl.SecurityIdentifier.Value;   //couldn't convert, return SID
                    }
                }
                if (type == DcomSecurityTypes.Launch)
                {
                    if ((dl.AccessMask & (int)MachineDComAccessRights.ExecuteLocal) != 0 ||
                        ((dl.AccessMask & (int)MachineDComAccessRights.Execute) != 0 &&
                         (dl.AccessMask & ((int)MachineDComAccessRights.ExecuteRemote |
                                           (int)MachineDComAccessRights.ActivateRemote |
                                           (int)MachineDComAccessRights.ActivateLocal)) == 0))
                    {
                        access.Add("LocalLaunch");
                    }
                    if ((dl.AccessMask & (int)MachineDComAccessRights.ExecuteRemote) != 0 ||
                        ((dl.AccessMask & (int)MachineDComAccessRights.Execute) != 0 &&
                         (dl.AccessMask & ((int)MachineDComAccessRights.ExecuteLocal |
                                           (int)MachineDComAccessRights.ActivateRemote |
                                           (int)MachineDComAccessRights.ActivateLocal)) == 0))
                    {
                        access.Add("RemoteLaunch");
                    }
                    if ((dl.AccessMask & (int)MachineDComAccessRights.ActivateLocal) != 0 ||
                        ((dl.AccessMask & (int)MachineDComAccessRights.Execute) != 0 &&
                         (dl.AccessMask & ((int)MachineDComAccessRights.ExecuteLocal |
                                           (int)MachineDComAccessRights.Execute |
                                           (int)MachineDComAccessRights.ActivateRemote)) == 0))
                    {
                        access.Add("LocalActivation");
                    }
                    if ((dl.AccessMask & (int)MachineDComAccessRights.ActivateRemote) != 0 ||
                        ((dl.AccessMask & (int)MachineDComAccessRights.Execute) != 0 &&
                         (dl.AccessMask & ((int)MachineDComAccessRights.ExecuteLocal |
                                           (int)MachineDComAccessRights.ExecuteRemote |
                                           (int)MachineDComAccessRights.ActivateLocal)) == 0))
                    {
                        access.Add("RemoteActivation");
                    }

                    dcomAce.User = user;
                    dcomAce.AccessType = accessControlType;
                    dcomAce.LocalAccess = null;
                    dcomAce.RemoteAccess = null;
                    dcomAce.LocalLaunch = access.Contains("LocalLaunch");
                    dcomAce.RemoteLaunch = access.Contains("RemoteLaunch");
                    dcomAce.LocalActivation = access.Contains("LocalActivation");
                    dcomAce.RemoteActivation = access.Contains("RemoteActivation");

                }
                else if (type == DcomSecurityTypes.Access)
                {
                    if ((dl.AccessMask & (int)MachineDComAccessRights.ExecuteLocal) != 0 ||
                        ((dl.AccessMask & (int)MachineDComAccessRights.Execute) != 0 &&
                         (dl.AccessMask & (int)MachineDComAccessRights.ExecuteRemote) == 0)) { access.Add("LocalAccess"); }
                    if ((dl.AccessMask & (int)MachineDComAccessRights.ExecuteRemote) != 0 ||
                        ((dl.AccessMask & (int)MachineDComAccessRights.Execute) != 0 &&
                         (dl.AccessMask & (int)MachineDComAccessRights.ExecuteLocal) == 0)) { access.Add("RemoteAccess"); }

                    dcomAce.User = user;
                    dcomAce.AccessType = accessControlType;
                    dcomAce.LocalAccess = access.Contains("LocalAccess");
                    dcomAce.RemoteAccess = access.Contains("RemoteAccess");
                    dcomAce.LocalLaunch = null;
                    dcomAce.RemoteLaunch = null;
                    dcomAce.LocalActivation = null;
                    dcomAce.RemoteActivation = null;

                }
                dcomAces.Add(dcomAce);
            }
            return dcomAces;
        }

        /// <summary>
        /// Grant an account the desired dcom privileges
        /// </summary>
        /// <param name="Type"></param>
        /// <param name="account"></param>
        /// <param name="rights"></param>
        /// <param name="option"></param>
        /// <param name="permissionType"></param>
        /// <param name="computerName"></param>
        /// <returns></returns>
        public static bool AddComputerRights(DcomSecurityTypes Type, SecurityIdentifier account, MachineDComAccessRights[] rights, 
            DcomPermissionOption option, AccessControlType permissionType, String computerName = null)
        {
            if (computerName == null)
            {
                computerName = String.Empty;
            }

            RawSecurityDescriptor rawSecurityDescriptor;
            MachineDComAccessRights[] testedRights = TestRights(rights, Type);
            int borInt;
            borInt = (int)MachineDComAccessRights.Execute;
            //Loop array, adding integer representation of right
            for (int x = 0; x < testedRights.Length; x++)
            {
                borInt |= (int)testedRights[x];
            }

            RegistryKey rootKey =
                RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, computerName);

            switch (Type)
            {
                case DcomSecurityTypes.Access:
                    if (option == DcomPermissionOption.Default)
                    {
                        byte[] regValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole")
                            .GetValue("DefaultAccessPermission");

                        if (regValue == null || regValue.Length == 0)
                        {
                            //Default value not set. Configure the default
                            //  Owner: BUILTIN\Administrators
                            //  Group            : BUILTIN\Administrators
                            //  DiscretionaryAcl : {
                            //      NT AUTHORITY\SELF: AccessAllowed(CreateDirectories, ListDirectory, WriteData), NT
                            //          AUTHORITY\SYSTEM: AccessAllowed(ListDirectory, WriteData), BUILTIN\Administrators: AccessAllowed
                            //          (CreateDirectories, ListDirectory, WriteData)}
                            //  SystemAcl: { }
                            RawSecurityDescriptor tempDescriptor = new RawSecurityDescriptor("O:BAG:BAD:(A;;CCDCLC;;;PS)(A;;CCDC;;;SY)(A;;CCDCLC;;;BA)");
                            byte[] tempBytes = new byte[tempDescriptor.BinaryLength];
                            tempDescriptor.GetBinaryForm(tempBytes, 0); //get binary array to temp var
                            regValue = tempBytes; //reset our main Var to byte array
                        }
                        rawSecurityDescriptor = new RawSecurityDescriptor(regValue, 0);
                        byte[] valBytes = NewBinarySecurityDescriptor(account, rawSecurityDescriptor,
                            permissionType, borInt);

                        rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole", true).SetValue("DefaultAccessPermission", valBytes, RegistryValueKind.Binary);

                        //Test Value and confirm
                        byte[] setValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole")
                            .GetValue("DefaultAccessPermission");

                        RawSecurityDescriptor fndDescriptor = new RawSecurityDescriptor(setValue, 0);
                        return ConfirmACE(account, fndDescriptor, permissionType, borInt);
                    }
                    else
                    {
                        byte[] regValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole")
                            .GetValue("MachineAccessRestriction");

                        rawSecurityDescriptor = new RawSecurityDescriptor(regValue, 0);
                        byte[] valBytes = NewBinarySecurityDescriptor(account, rawSecurityDescriptor,
                            permissionType, borInt);
                        rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole", true).SetValue("MachineAccessRestriction", valBytes, RegistryValueKind.Binary);

                        byte[] setValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole")
                            .GetValue("MachineAccessRestriction");

                        RawSecurityDescriptor fndDescriptor = new RawSecurityDescriptor(setValue, 0);
                        return ConfirmACE(account, fndDescriptor, permissionType, borInt);
                    }
                case DcomSecurityTypes.Launch:
                    if (option == DcomPermissionOption.Default)
                    {
                        byte[] regValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.ChangePermissions)
                            .GetValue("DefaultLaunchPermission");

                        rawSecurityDescriptor = new RawSecurityDescriptor(regValue, 0);
                        byte[] valBytes = NewBinarySecurityDescriptor(account, rawSecurityDescriptor,
                            permissionType, borInt);
                        rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole", true).SetValue("DefaultLaunchPermission", valBytes, RegistryValueKind.Binary);

                        byte[] setValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole")
                            .GetValue("DefaultLaunchPermission");


                        RawSecurityDescriptor fndDescriptor = new RawSecurityDescriptor(setValue, 0);
                        return ConfirmACE(account, fndDescriptor, permissionType, borInt);
                    }
                    else
                    {
                        byte[] regValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole")
                            .GetValue("MachineLaunchRestriction");

                        rawSecurityDescriptor = new RawSecurityDescriptor(regValue, 0);
                        byte[] valBytes = NewBinarySecurityDescriptor(account, rawSecurityDescriptor,
                            permissionType, borInt);
                        rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole", true).SetValue("MachineLaunchRestriction", valBytes, RegistryValueKind.Binary);

                        byte[] setValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole")
                            .GetValue("MachineLaunchRestriction");

                        RawSecurityDescriptor fndDescriptor = new RawSecurityDescriptor(setValue, 0);
                        return ConfirmACE(account, fndDescriptor, permissionType, borInt);
                    }
            }
            return false;
        }
        /// <summary>
        /// Grant an account the desired privileges
        /// </summary>
        /// <param name="Type"></param>
        /// <param name="account"></param>
        /// <param name="rights"></param>
        /// <param name="permissionType"></param>
        /// <param name="appId"></param>
        /// <param name="computerName"></param>
        /// <param name="revoke">Instead of adding or updating rights for the <paramref name="account"/>, this will remove the account from the Access Control List.</param>
        /// <returns></returns>
        public static bool AddApplicationRights(DcomSecurityTypes Type, SecurityIdentifier account, MachineDComAccessRights[] rights, AccessControlType permissionType, string appId, 
            String computerName = null, bool revoke = false)
        {
            string regPath = "AppId\\" + appId;
            if (computerName == null)
            {
                computerName = String.Empty;
            }
            try
            {
                byte[] regValue = null;
                RawSecurityDescriptor rawSecurityDescriptor;

                MachineDComAccessRights[] testedRights = TestRights(rights, Type);
                int borInt;
                borInt = (int)MachineDComAccessRights.Execute;
                //Loop array, adding integer representation of right
                for (int x = 0; x < testedRights.Length; x++)
                {
                    borInt |= (int)testedRights[x];
                }
                byte[] valBytes = null;
                byte[] setValue = null;
                RawSecurityDescriptor fndDescriptor;
                RegistryKey rootKey =
                    RegistryKey.OpenRemoteBaseKey(RegistryHive.ClassesRoot, computerName);

                switch (Type)
                {
                    case DcomSecurityTypes.Access:
                        regValue = (byte[])rootKey.OpenSubKey(regPath)
                            .GetValue("AccessPermission");

                        if (regValue == null || regValue.Length == 0)
                        {
                            //Default value not set. Configure the default
                            //  Owner: BUILTIN\Administrators
                            //  Group            : BUILTIN\Administrators
                            //  DiscretionaryAcl : {
                            //      NT AUTHORITY\SELF: AccessAllowed(CreateDirectories, ListDirectory, WriteData), NT
                            //          AUTHORITY\SYSTEM: AccessAllowed(ListDirectory, WriteData), BUILTIN\Administrators: AccessAllowed
                            //          (CreateDirectories, ListDirectory, WriteData)}
                            //  SystemAcl: { }
                            RawSecurityDescriptor tempDescriptor = new RawSecurityDescriptor("O:BAG:BAD:(A;;CCDCLC;;;PS)(A;;CCDC;;;SY)(A;;CCDCLC;;;BA)");
                            byte[] tempBytes = new byte[tempDescriptor.BinaryLength];
                            tempDescriptor.GetBinaryForm(tempBytes, 0); //get binary array to temp var
                            regValue = tempBytes; //reset our main Var to byte array
                        }
                        rawSecurityDescriptor = new RawSecurityDescriptor(regValue, 0);
                        valBytes = NewBinarySecurityDescriptor(account, rawSecurityDescriptor,
                            permissionType, borInt, revoke);
                        rootKey.OpenSubKey(regPath, true).SetValue("AccessPermission", valBytes, RegistryValueKind.Binary);

                        setValue = (byte[])rootKey.OpenSubKey(regPath)
                            .GetValue("AccessPermission");

                        fndDescriptor = new RawSecurityDescriptor(setValue, 0);
                        return ConfirmACE(account, fndDescriptor, permissionType, borInt);
                    case DcomSecurityTypes.Launch:
                        regValue = (byte[])rootKey.OpenSubKey(regPath)
                            .GetValue("LaunchPermission");

                        if (regValue == null || regValue.Length == 0)
                        {
                            //Default value not set. Configure the default
                            //  Owner: BUILTIN\Administrators
                            //  Group            : BUILTIN\Administrators
                            //  DiscretionaryAcl : {
                            //      NT AUTHORITY\SELF: AccessAllowed(CreateDirectories, ListDirectory, WriteData), NT
                            //          AUTHORITY\SYSTEM: AccessAllowed(ListDirectory, WriteData), BUILTIN\Administrators: AccessAllowed
                            //          (CreateDirectories, ListDirectory, WriteData)}
                            //  SystemAcl: { }
                            RawSecurityDescriptor tempDescriptor = new RawSecurityDescriptor("O:BAG:BAD:(A;;CCDCLC;;;PS)(A;;CCDC;;;SY)(A;;CCDCLC;;;BA)");
                            byte[] tempBytes = new byte[tempDescriptor.BinaryLength];
                            tempDescriptor.GetBinaryForm(tempBytes, 0); //get binary array to temp var
                            regValue = tempBytes; //reset our main Var to byte array
                        }
                        rawSecurityDescriptor = new RawSecurityDescriptor(regValue, 0);
                        valBytes = NewBinarySecurityDescriptor(account, rawSecurityDescriptor,
                            permissionType, borInt, revoke);

                        rootKey.OpenSubKey(regPath, true).SetValue("LaunchPermission", valBytes, RegistryValueKind.Binary);

                        setValue = (byte[])rootKey.OpenSubKey(regPath)
                            .GetValue("LaunchPermission");

                        fndDescriptor = new RawSecurityDescriptor(setValue, 0);
                        return ConfirmACE(account, fndDescriptor, permissionType, borInt);
                }
            }
            catch (ArgumentNullException arg)
            {
                throw new UnauthorizedAccessException("The registry key security setting is null", arg);
            }
            catch (SecurityException sec)
            {
                throw new UnauthorizedAccessException("The user does not have the permissions required to create or modify registry keys.", sec);
            }
            catch (UnauthorizedAccessException unauthorized)
            {
                throw new UnauthorizedAccessException(
                    $"The registry key is read-only, or has not been opened with admin or write access. Ensure application calling application is 'Run as Administrator'. Registry Key : {regPath}", unauthorized);
            }
            return false;
        }
        /// <summary>
        /// Revoke the given account privileges
        /// </summary>
        /// <param name="Type"></param>
        /// <param name="account"></param>
        /// <param name="option"></param>
        /// <param name="computerName"></param>
        /// <returns></returns>
        public static bool RemoveComputerRights(DcomSecurityTypes Type, SecurityIdentifier account, DcomPermissionOption option, String computerName = null)
        {
            if (computerName == null)
            {
                computerName = String.Empty;
            }

            RawSecurityDescriptor rawSecurityDescriptor;

            RegistryKey rootKey =
                RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, computerName);

            switch (Type)
            {
                case DcomSecurityTypes.Access:
                    if (option == DcomPermissionOption.Default)
                    {
                        byte[] regValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole")
                            .GetValue("DefaultAccessPermission");

                        if (regValue == null || regValue.Length == 0)
                        {
                            //Default value not set. Configure the default
                            RawSecurityDescriptor tempDescriptor = new RawSecurityDescriptor("O:BAG:BAD:(A;;CCDCLC;;;PS)(A;;CCDC;;;SY)(A;;CCDCLC;;;BA)");
                            byte[] tempBytes = new byte[tempDescriptor.BinaryLength];
                            tempDescriptor.GetBinaryForm(tempBytes, 0); //get binary array to temp var
                            regValue = tempBytes; //reset our main Var to byte array
                        }
                        rawSecurityDescriptor = new RawSecurityDescriptor(regValue, 0);
                        byte[] valBytes = NewBinarySecurityDescriptor(account, rawSecurityDescriptor,
                            AccessControlType.Deny, 0, true);

                        rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole", true).SetValue("DefaultAccessPermission", valBytes, RegistryValueKind.Binary);

                        //Test Value and confirm
                        byte[] setValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole")
                            .GetValue("DefaultAccessPermission");

                        RawSecurityDescriptor fndDescriptor = new RawSecurityDescriptor(setValue, 0);
                        return ConfirmACERemoval(account, fndDescriptor);
                    }
                    else
                    {
                        byte[] regValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole")
                            .GetValue("MachineAccessRestriction");

                        rawSecurityDescriptor = new RawSecurityDescriptor(regValue, 0);
                        byte[] valBytes = NewBinarySecurityDescriptor(account, rawSecurityDescriptor,
                            AccessControlType.Deny, 0, true);
                        rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole", true).SetValue("MachineAccessRestriction", valBytes, RegistryValueKind.Binary);

                        //Test Value and confirm
                        byte[] setValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole")
                            .GetValue("MachineAccessRestriction");

                        RawSecurityDescriptor fndDescriptor = new RawSecurityDescriptor(setValue, 0);
                        return ConfirmACERemoval(account, fndDescriptor);
                    }
                case DcomSecurityTypes.Launch:
                    if (option == DcomPermissionOption.Default)
                    {
                        byte[] regValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.ChangePermissions)
                            .GetValue("DefaultLaunchPermission");

                        rawSecurityDescriptor = new RawSecurityDescriptor(regValue, 0);
                        byte[] valBytes = NewBinarySecurityDescriptor(account, rawSecurityDescriptor,
                            AccessControlType.Deny, 0, true);
                        rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole", true).SetValue("DefaultLaunchPermission", valBytes, RegistryValueKind.Binary);

                        //Test Value and confirm
                        byte[] setValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole")
                            .GetValue("DefaultLaunchPermission");


                        RawSecurityDescriptor fndDescriptor = new RawSecurityDescriptor(setValue, 0);
                        return ConfirmACERemoval(account, fndDescriptor);
                    }
                    else
                    {
                        byte[] regValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole")
                            .GetValue("MachineLaunchRestriction");

                        rawSecurityDescriptor = new RawSecurityDescriptor(regValue, 0);
                        byte[] valBytes = NewBinarySecurityDescriptor(account, rawSecurityDescriptor,
                            AccessControlType.Deny, 0, true);
                        rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole", true).SetValue("MachineLaunchRestriction", valBytes, RegistryValueKind.Binary);

                        //Test Value and confirm
                        byte[] setValue = (byte[])rootKey.OpenSubKey("SOFTWARE\\Microsoft\\Ole")
                            .GetValue("MachineLaunchRestriction");

                        RawSecurityDescriptor fndDescriptor = new RawSecurityDescriptor(setValue, 0);
                        return ConfirmACERemoval(account, fndDescriptor);
                    }

            }
            return false;
        }

        /// <summary>
        /// Revoke the accounts given privileges
        /// </summary>
        /// <param name="Type"></param>
        /// <param name="account"></param>
        /// <param name="appId"></param>
        /// <param name="computerName"></param>
        /// <returns></returns>
        public static bool RemoveApplicationRights(DcomSecurityTypes Type, SecurityIdentifier account, string appId, String computerName = null)
        {
            string regPath = "AppId\\" + appId;
            if (computerName == null)
            {
                computerName = String.Empty;
            }
            try
            {
                byte[] regValue = null;
                RawSecurityDescriptor rawSecurityDescriptor;

                byte[] valBytes = null;
                byte[] setValue = null;
                RawSecurityDescriptor fndDescriptor;
                RegistryKey rootKey =
                    RegistryKey.OpenRemoteBaseKey(RegistryHive.ClassesRoot, computerName);

                switch (Type)
                {
                    case DcomSecurityTypes.Access:
                        regValue = (byte[])rootKey.OpenSubKey(regPath)
                            .GetValue("AccessPermission");

                        if (regValue == null || regValue.Length == 0)
                        {
                            //Default value not set. Configure the default
                            //  Owner: BUILTIN\Administrators
                            //  Group            : BUILTIN\Administrators
                            //  DiscretionaryAcl : {
                            //      NT AUTHORITY\SELF: AccessAllowed(CreateDirectories, ListDirectory, WriteData), NT
                            //          AUTHORITY\SYSTEM: AccessAllowed(ListDirectory, WriteData), BUILTIN\Administrators: AccessAllowed
                            //          (CreateDirectories, ListDirectory, WriteData)}
                            //  SystemAcl: { }
                            RawSecurityDescriptor tempDescriptor = new RawSecurityDescriptor("O:BAG:BAD:(A;;CCDCLC;;;PS)(A;;CCDC;;;SY)(A;;CCDCLC;;;BA)");
                            byte[] tempBytes = new byte[tempDescriptor.BinaryLength];
                            tempDescriptor.GetBinaryForm(tempBytes, 0); //get binary array to temp var
                            regValue = tempBytes; //reset our main Var to byte array
                        }
                        rawSecurityDescriptor = new RawSecurityDescriptor(regValue, 0);
                        valBytes = NewBinarySecurityDescriptor(account, rawSecurityDescriptor,
                            AccessControlType.Deny, 0, true);

                        rootKey.OpenSubKey(regPath, true).SetValue("AccessPermission", valBytes, RegistryValueKind.Binary);

                        //Test Value and confirm
                        setValue = (byte[])rootKey.OpenSubKey(regPath)
                            .GetValue("AccessPermission");

                        fndDescriptor = new RawSecurityDescriptor(setValue, 0);
                        return ConfirmACERemoval(account, fndDescriptor);
                    case DcomSecurityTypes.Launch:
                        regValue = (byte[])rootKey.OpenSubKey(regPath)
                            .GetValue("LaunchPermission");

                        if (regValue == null || regValue.Length == 0)
                        {
                            //Default value not set. Configure the default
                            //  Owner: BUILTIN\Administrators
                            //  Group            : BUILTIN\Administrators
                            //  DiscretionaryAcl : {
                            //      NT AUTHORITY\SELF: AccessAllowed(CreateDirectories, ListDirectory, WriteData), NT
                            //          AUTHORITY\SYSTEM: AccessAllowed(ListDirectory, WriteData), BUILTIN\Administrators: AccessAllowed
                            //          (CreateDirectories, ListDirectory, WriteData)}
                            //  SystemAcl: { }
                            RawSecurityDescriptor tempDescriptor = new RawSecurityDescriptor("O:BAG:BAD:(A;;CCDCLC;;;PS)(A;;CCDC;;;SY)(A;;CCDCLC;;;BA)");
                            byte[] tempBytes = new byte[tempDescriptor.BinaryLength];
                            tempDescriptor.GetBinaryForm(tempBytes, 0); //get binary array to temp var
                            regValue = tempBytes; //reset our main Var to byte array
                        }
                        rawSecurityDescriptor = new RawSecurityDescriptor(regValue, 0);
                        valBytes = NewBinarySecurityDescriptor(account, rawSecurityDescriptor,
                            AccessControlType.Deny, 0, true);

                        rootKey.OpenSubKey(regPath, true).SetValue("LaunchPermission", valBytes, RegistryValueKind.Binary);

                        //Test Value and confirm
                        setValue = (byte[])rootKey.OpenSubKey(regPath)
                            .GetValue("LaunchPermission");

                        fndDescriptor = new RawSecurityDescriptor(setValue, 0);
                        return ConfirmACERemoval(account, fndDescriptor);
                }
            }
            catch (ArgumentNullException arg)
            {
                throw new UnauthorizedAccessException("The registry key security setting is null", arg);
            }
            catch (SecurityException sec)
            {
                throw new UnauthorizedAccessException("The user does not have the permissions required to create or modify registry keys.", sec);
            }
            catch (UnauthorizedAccessException unauthorized)
            {
                throw new UnauthorizedAccessException(
                    $"The registry key is read-only, or has not been opened with admin or write access. Ensure application calling application is 'Run as Administrator'. Registry Key : {regPath}", unauthorized);
            }
            return false;
        }

        /// <summary>
        /// Reset the application to use Global Default settings, rather than custom settings
        /// </summary>
        /// <param name="Type"></param>
        /// <param name="appId"></param>
        /// <param name="computerName"></param>
        /// <returns></returns>
        public static bool ResetApplicationRightsToDefault(DcomSecurityTypes Type, string appId, String computerName = null)
        {
            string regPath = "AppId\\" + appId;
            if (computerName == null)
            {
                computerName = String.Empty;
            }
            try
            {
                byte[] regValue = null;
                
                RegistryKey rootKey =
                    RegistryKey.OpenRemoteBaseKey(RegistryHive.ClassesRoot, computerName);

                switch (Type)
                {
                    case DcomSecurityTypes.Access:
                        regValue = (byte[])rootKey.OpenSubKey(regPath)
                            .GetValue("AccessPermission");
                        if (regValue != null || regValue.Length > 0)
                        {
                            rootKey.OpenSubKey(regPath).DeleteValue("AccessPermission");
                        }
                        return true;
                    case DcomSecurityTypes.Launch:
                        regValue = (byte[])rootKey.OpenSubKey(regPath)
                            .GetValue("LaunchPermission");
                        if (regValue != null || regValue.Length > 0)
                        {
                            rootKey.OpenSubKey(regPath).DeleteValue("LaunchPermission");
                        }
                        return true;
                }
            }
            catch (SecurityException sec)
            {
                throw new UnauthorizedAccessException("The user does not have the permissions required to create or modify registry keys.", sec);
            }
            catch (UnauthorizedAccessException unauthorized)
            {
                throw new UnauthorizedAccessException(
                    $"The registry key is read-only, or has not been opened with admin or write access. Ensure application calling application is 'Run as Administrator'. Registry Key : {regPath}", unauthorized);
            }
            return false;
        }

        private static byte[] NewBinarySecurityDescriptor(SecurityIdentifier account, RawSecurityDescriptor rawSecurityDescriptor, 
            AccessControlType permissionType, int accessRight, bool revoke = false)
        {
            if (account == null)
            {
                throw new ArgumentNullException(nameof(account));
            }
            try
            {
                RawAcl doc;
                int i = 0;
                if (rawSecurityDescriptor.DiscretionaryAcl == null)
                {
                    //Default value not set. Configure the default
                    //  Owner: BUILTIN\Administrators
                    //  Group            : BUILTIN\Administrators
                    //  DiscretionaryAcl : {
                    //      NT AUTHORITY\SELF: AccessAllowed(CreateDirectories, ListDirectory, WriteData), NT
                    //          AUTHORITY\SYSTEM: AccessAllowed(ListDirectory, WriteData), BUILTIN\Administrators: AccessAllowed
                    //          (CreateDirectories, ListDirectory, WriteData)}
                    //  SystemAcl: { }
                    RawSecurityDescriptor temp = new RawSecurityDescriptor("O:BAG:BAD:(A;;CCDCLC;;;PS)(A;;CCDC;;;SY)(A;;CCDCLC;;;BA)");
                    doc = new RawAcl(temp.DiscretionaryAcl.Revision, 1);
                }
                else
                {
                    doc = rawSecurityDescriptor.DiscretionaryAcl;
                }

                if (doc == null)
                {
                    throw new ArgumentNullException("Failed to generate Acl");
                }
                //If doc is NULL, then creating a brand new ACL, and trying to loop it will create exception
                foreach (var genericAce in doc)
                {
                    var ca = (CommonAce)genericAce;
                    if (ca.SecurityIdentifier.Equals(account))
                    {
                        //found a matching SID in pre-existing Rights list
                        if (!revoke && (AccessControlType)ca.AceType == permissionType)
                        {
                            //We've a pre-existing rule configured
                            //Remove it, so we can replace with new defined rule
                            doc.RemoveAce(i);
                        }
                        if (revoke)
                        {
                            doc.RemoveAce(i);
                        }
                    }
                    i += 1;
                }

                //add new ACE now
                // A canonical ACL must have ACES sorted according to the following order:
                //   1. Access-denied on the object
                //   2. Access-denied on a child or property
                //   3. Access-allowed on the object
                //   4. Access-allowed on a child or property
                //   5. All inherited ACEs

                if (!revoke)
                {
                    CommonAce newAce = new CommonAce(AceFlags.None, (AceQualifier)permissionType, accessRight, account, false, null);
                    doc.InsertAce(doc.Count, newAce);
                }

                var implicitDenyDacl = new List<CommonAce>();
                var implicitDenyObjectDacl = new List<CommonAce>();
                var inheritedDacl = new List<CommonAce>();
                var implicitAllowDacl = new List<CommonAce>();
                var implicitAllowObjectDacl = new List<CommonAce>();

                foreach (var genericAce in doc)
                {
                    var ace = (CommonAce)genericAce;
                    if ((ace.AceFlags & AceFlags.Inherited) == AceFlags.Inherited)
                    {
                        inheritedDacl.Add(ace);
                    }
                    else
                    {
                        switch (ace.AceType)
                        {
                            case AceType.AccessAllowed:

                                implicitAllowDacl.Add(ace);
                                break;
                            case AceType.AccessDenied:
                                implicitDenyDacl.Add(ace);
                                break;
                            case AceType.AccessAllowedObject:
                                implicitAllowObjectDacl.Add(ace);
                                break;
                            case AceType.AccessDeniedObject:
                                implicitDenyObjectDacl.Add(ace);
                                break;
                        }
                    }
                }
                Int32 aceIndex = 0;

                RawAcl newDacl = new RawAcl(doc.Revision, doc.Count);
                implicitDenyDacl.ForEach(x => newDacl.InsertAce(aceIndex++, x));
                implicitDenyObjectDacl.ForEach(x => newDacl.InsertAce(aceIndex++, x));
                implicitAllowDacl.ForEach(x => newDacl.InsertAce(aceIndex++, x));
                implicitAllowObjectDacl.ForEach(x => newDacl.InsertAce(aceIndex++, x));
                inheritedDacl.ForEach(x => newDacl.InsertAce(aceIndex++, x));

                if (aceIndex != doc.Count)
                {
                    throw new Exception("The DACL cannot be canonicalized since it would potentially result in a loss of information");
                }

                rawSecurityDescriptor.DiscretionaryAcl = newDacl;

                byte[] binaryform = new byte[rawSecurityDescriptor.BinaryLength];
                rawSecurityDescriptor.GetBinaryForm(binaryform, 0);
                if (binaryform.Length == 0) { throw new SecurityException(
                    $"Failed to generate Binary Security Descriptor for Account ({account.Translate(typeof(NTAccount)).Value})"); }
                return binaryform;
            }
            catch (Exception ex)
            {
                throw new ArgumentException(
                    $"Unable to generate Binary Security Descriptor for Account ({account.Translate(typeof(NTAccount)).Value})", ex);
            }
        }

        private static bool ConfirmACE(SecurityIdentifier account, RawSecurityDescriptor rawSecurityDescriptor, AccessControlType permissionType, int accessRight)
        {
            try
            {
                bool found = false;
                foreach (var genericAce in rawSecurityDescriptor.DiscretionaryAcl)
                {
                    var ca = (CommonAce)genericAce;
                    if (ca.SecurityIdentifier.Equals(account) && (AccessControlType)ca.AceType == permissionType && ca.AccessMask == accessRight)
                    {
                        //found a matching SID in pre-existing Rights list
                        //We've a pre-existing rule configured
                        //Remove it, so we can replace with new defined rule
                        found = true;
                    }
                }
                return found;
            }
            catch (Exception ex)
            {
                throw new ArgumentException(
                    $"Unable to find Binary Security Descriptor for Account ({account.Translate(typeof(NTAccount)).Value})", ex);
            }
        }

        private static bool ConfirmACERemoval(SecurityIdentifier account, RawSecurityDescriptor rawSecurityDescriptor)
        {
            try
            {
                bool success = true;
                foreach (var genericAce in rawSecurityDescriptor.DiscretionaryAcl)
                {
                    var ca = (CommonAce)genericAce;
                    if (ca.SecurityIdentifier.Equals(account))
                    {
                        //found a matching SID in pre-existing Rights list
                        success = false;
                    }
                }
                return success;
            }
            catch (Exception ex)
            {
                throw new ArgumentException(
                    $"Unable to find Binary Security Descriptor for Account ({account.Translate(typeof(NTAccount)).Value})", ex);
            }
        }

        private static MachineDComAccessRights[] TestRights(MachineDComAccessRights[] Rights, DcomSecurityTypes option)
        {
            var lst = new List<MachineDComAccessRights>(Rights);
            switch (option)
            {
                case DcomSecurityTypes.Launch:
                    if (lst.Contains(MachineDComAccessRights.ActivateLocal) ||
                        lst.Contains(MachineDComAccessRights.ActivateRemote))
                    {
                        // todo complete removal of bad rights
                        var cnt = lst.RemoveAll(m =>
                            m.Equals(MachineDComAccessRights.ActivateLocal) ||
                            m.Equals(MachineDComAccessRights.ActivateRemote));
                        //MachineDComAccessRights[] ret = lst.SkipWhile(element =>
                        //    !element.Equals(MachineDComAccessRights.ActivateLocal) &&
                        //    !element.Equals(MachineDComAccessRights.ActivateRemote)).ToArray();
                        return lst.ToArray();
                    }
                    return Rights;
            }

            return Rights;
        }
        enum SID_NAME_USE
        {
            SidTypeUser = 1,
            SidTypeGroup,
            SidTypeDomain,
            SidTypeAlias,
            SidTypeWellKnownGroup,
            SidTypeDeletedAccount,
            SidTypeInvalid,
            SidTypeUnknown,
            SidTypeComputer
        }

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool LookupAccountSid(
            string lpSystemName,
            [MarshalAs(UnmanagedType.LPArray)] byte[] Sid,
            StringBuilder lpName,
            ref uint cchName,
            StringBuilder ReferencedDomainName,
            ref uint cchReferencedDomainName,
            out SID_NAME_USE peUse);

    }
}