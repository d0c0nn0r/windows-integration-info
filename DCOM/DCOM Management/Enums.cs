/*************************************************************************
* 
* TQS INTEGRATION CONFIDENTIAL
* __________________
* 
*  [2006] - [2018] TQS Integration LLC 
*  All Rights Reserved.
* 
* NOTICE:  All information contained herein is, and remains
* the property of TQS Integration LLC and its suppliers,
* if any.  The intellectual and technical concepts contained
* herein are proprietary to TQS Integration LLC
* and its suppliers and may be covered by U.S. and Foreign Patents,
* patents in process, and are protected by trade secret or copyright law.
* Dissemination of this information or reproduction of this material
* is strictly forbidden unless prior written permission is obtained
* from TQS Integration LLC.
*/ 

using System;
//using System.ComponentModel;

namespace WinDevOps
{
    /// <summary>
    /// The options to configure a domain account with delegation settings.
    /// </summary>
    [Flags]
    public enum DelegationConfig
    {
        /// <summary>
        /// No delegation is possible
        /// </summary>
        None =0,
        /// <summary>
        /// Do not trust this user for delegation
        /// </summary>
        DoNotTrustForDelegation=1,
        /// <summary>
        /// Trust this user for delegation to any service (Kerberos only)
        /// </summary>
        TrustForDelegationToAnyService_KerberosOnly = 2,
        /// <summary>
        /// Trust this user for delegation to specified services only. Use Kerberos Only
        /// </summary>
        TrustForDelegationToSpecifiedService_KerberosOnly = 4,
        /// <summary>
        /// Trust this user for delegation to specified services only. Use any authentication protocol
        /// </summary>
        TrustForDelegationToSpecifiedService_AnyAuthProtocol = 8
    }

    /// <summary>
    /// Flags that control the behavior of the user account.
    /// </summary>
    [Flags]
    public enum UserAccountControl
    {
        /// <summary>
        /// The logon script is executed. 
        ///</summary>
        SCRIPT = 0x00000001,

        /// <summary>
        /// The user account is disabled. 
        ///</summary>
        ACCOUNTDISABLE = 0x00000002,

        /// <summary>
        /// The home directory is required. 
        ///</summary>
        HOMEDIR_REQUIRED = 0x00000008,

        /// <summary>
        /// The account is currently locked out. 
        ///</summary>
        LOCKOUT = 0x00000010,

        /// <summary>
        /// No password is required. 
        ///</summary>
        PASSWD_NOTREQD = 0x00000020,

        /// <summary>
        /// The user cannot change the password. 
        ///</summary>
        /// <remarks>Note:  You cannot assign the permission settings of PASSWD_CANT_CHANGE by directly modifying the UserAccountControl attribute. </remarks>
        /// <remarks>For more information and a code example that shows how to prevent a user from changing the password, see User Cannot Change Password.</remarks>
        PASSWD_CANT_CHANGE = 0x00000040,

        /// <summary>
        /// The user can send an encrypted password. 
        ///</summary>
        ENCRYPTED_TEXT_PASSWORD_ALLOWED = 0x00000080,

        /// <summary>
        /// This is an account for users whose primary account is in another domain. This account provides user access to this domain, but not 
        /// to any domain that trusts this domain. Also known as a local user account. 
        ///</summary>
        TEMP_DUPLICATE_ACCOUNT = 0x00000100,

        /// <summary>
        /// This is a default account type that represents a typical user. 
        ///</summary>
        NORMAL_ACCOUNT = 0x00000200,

        /// <summary>
        /// This is a permit to trust account for a system domain that trusts other domains. 
        ///</summary>
        INTERDOMAIN_TRUST_ACCOUNT = 0x00000800,

        /// <summary>
        /// This is a computer account for a computer that is a member of this domain. 
        ///</summary>
        WORKSTATION_TRUST_ACCOUNT = 0x00001000,

        /// <summary>
        /// This is a computer account for a system backup domain controller that is a member of this domain. 
        ///</summary>
        SERVER_TRUST_ACCOUNT = 0x00002000,

        /// <summary>
        /// Not used. 
        ///</summary>
        Unused1 = 0x00004000,

        /// <summary>
        /// Not used. 
        ///</summary>
        Unused2 = 0x00008000,

        /// <summary>
        /// The password for this account will never expire. 
        ///</summary>
        DONT_EXPIRE_PASSWD = 0x00010000,

        /// <summary>
        /// This is an MNS logon account. 
        ///</summary>
        MNS_LOGON_ACCOUNT = 0x00020000,

        /// <summary>
        /// The user must log on using a smart card. 
        ///</summary>
        SMARTCARD_REQUIRED = 0x00040000,

        /// <summary>
        /// The service account (user or computer account), under which a service runs, is trusted for Kerberos delegation. Any such service 
        /// can impersonate a client requesting the service. 
        ///</summary>
        TRUSTED_FOR_DELEGATION = 0x00080000,

        /// <summary>
        /// The security context of the user will not be delegated to a service even if the service account is set as trusted for Kerberos delegation. 
        ///</summary>
        NOT_DELEGATED = 0x00100000,

        /// <summary>
        /// Restrict this principal to use only Data Encryption Standard (DES) encryption types for keys. 
        ///</summary>
        USE_DES_KEY_ONLY = 0x00200000,

        /// <summary>
        /// This account does not require Kerberos pre-authentication for logon. 
        ///</summary>
        DONT_REQUIRE_PREAUTH = 0x00400000,

        /// <summary>
        /// The user password has expired. This flag is created by the system using data from the Pwd-Last-Set attribute and the domain policy. 
        ///</summary>
        PASSWORD_EXPIRED = 0x00800000,

        /// <summary>
        /// The account is enabled for delegation. This is a security-sensitive setting; accounts with this option enabled should be strictly 
        /// controlled. This setting enables a service running under the account to assume a client identity and authenticate as that user to 
        /// other remote servers on the network.
        ///</summary>
        TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 0x01000000,

        /// <summary>
        /// 
        /// </summary>
        PARTIAL_SECRETS_ACCOUNT = 0x04000000,

        /// <summary>
        /// 
        /// </summary>
        USE_AES_KEYS = 0x08000000
    }

    /// <summary>
    /// The type of account, as stored in the Active Directory
    /// </summary>
    public enum sAMAccountType
    {
        /// <summary>
        /// Domain Object
        /// </summary>
        SAM_DOMAIN_OBJECT = 0x0,
        /// <summary>
        /// Group Object
        /// </summary>
        SAM_GROUP_OBJECT = 0x10000000,
        /// <summary>
        /// Non-Security Group Object
        /// </summary>
        SAM_NON_SECURITY_GROUP_OBJECT = 0x10000001,
        /// <summary>
        /// Alias Object
        /// </summary>
        SAM_ALIAS_OBJECT = 0x20000000,
        /// <summary>
        /// Non-Security Alias Object
        /// </summary>
        SAM_NON_SECURITY_ALIAS_OBJECT = 0x20000001,
        /// <summary>
        /// User Object
        /// </summary>
        SAM_USER_OBJECT = 0x30000000,
        /// <summary>
        /// Machine Object
        /// </summary>
        SAM_MACHINE_ACCOUNT = 0x30000001,
        /// <summary>
        /// Trusted account Object
        /// </summary>
        SAM_TRUST_ACCOUNT = 0x30000002,
        /// <summary>
        /// App Basic Group Object
        /// </summary>
        SAM_APP_BASIC_GROUP = 0x40000000,
        /// <summary>
        /// App Query Group Object
        /// </summary>
        SAM_APP_QUERY_GROUP = 0x40000001
    }

    /// <summary>
    /// Flags that indicate the Authentication Level of a COM Application
    /// https://msdn.microsoft.com/en-us/library/windows/desktop/ms679675(v=vs.85).aspx
    /// </summary>
    [Flags]
    public enum RPC_C_AUTHN
    {
        /// <summary>
        /// RPC_C_AUTHN_LEVEL_DEFAULT
        /// </summary>
        Default = 0,
        /// <summary>
        /// RPC_C_AUTHN_LEVEL_NONE
        /// </summary>
        None = 1,
        /// <summary>
        /// RPC_C_AUTHN_LEVEL_CONNECT
        /// </summary>
        Connect = 2,
        /// <summary>
        /// RPC_C_AUTHN_LEVEL_CALL
        /// </summary>
        Call = 3,
        /// <summary>
        /// RPC_C_AUTHN_LEVEL_PKT
        /// </summary>
        Packet = 4,
        /// <summary>
        /// RPC_C_AUTHN_LEVEL_PKT_INTEGRITY
        /// </summary>
        Packet_Integrity = 5,
        /// <summary>
        /// RPC_C_AUTHN_LEVEL_PKT_PRIVACY
        /// </summary>
        Packet_Privacy = 6
    }

    /// <summary>
    /// The Access-Control-List options for DCOM settings
    /// </summary>

    public enum DcomSecurityTypes
    {
        /// <summary>
        /// Describes the Access Control List (ACL) of the principals that can access instances of this class. This ACL is used only by applications that do not call CoInitializeSecurity.
        /// </summary>
        Access,
        /// <summary>
        /// Describes the Access Control List (ACL) of the principals that can start new servers for this class.
        /// </summary>
        Launch,
        /// <summary>
        /// Describes the Access Control List (ACL) of the registry key that maintains the DCOM configuration.
        /// </summary>
        Config
    }
    /// <summary>
    /// The permission collection options for DCOM settings
    /// </summary>
    public enum DcomPermissionOption
    {
        /// <summary>
        /// Is not applicable
        /// </summary>
        None,
        /// <summary>
        /// Limits option
        /// </summary>
        Limits,
        /// <summary>
        /// Default option
        /// </summary>
        Default
    }

    /// <summary>
    /// The DCOM rights that can be assigned in an Access-Control-List
    /// </summary>

    [Flags]
    public enum MachineDComAccessRights : uint
    {
        /// <summary>
        /// Execute permission
        /// </summary>
        Execute = 1,
        /// <summary>
        /// Execute-Local permission
        /// </summary>
        ExecuteLocal = 2,
        /// <summary>
        /// Execute-remote permission
        /// </summary>
        ExecuteRemote = 4,
        /// <summary>
        /// Activate-Local permission
        /// </summary>
        ActivateLocal = 8,
        /// <summary>
        /// Activate-Remote permission
        /// </summary>
        ActivateRemote = 16
    }
    /// <summary>
    /// The startup options for a DCOM-supporting service
    /// </summary>
    [Flags]
    public enum ServiceStart
    {
        /// <summary>
        /// The kernel loaded will load this driver first as its needed to use the boot volume device
        /// </summary>
        Boot = 0,
        /// <summary>
        /// This is loaded by the I/O subsystem
        /// </summary>
        System = 1,
        /// <summary>
        /// The service is always loaded and run
        /// </summary>
        Automatic = 2,
        /// <summary>
        /// This service does not start automatically and must be manually started by the user
        /// </summary>
        Manual = 3,
        /// <summary>
        /// The service is disabled and should not be started
        /// </summary>
        Disabled = 4
    }

    /// <summary>
    /// Privileged rights to be able to add/remove from Windows Application Servers.
    /// </summary>
    public enum Rights
    {
        /// <summary>
        /// Access Credential Manager as a trusted caller
        /// </summary>
        SeTrustedCredManAccessPrivilege,
        /// <summary>
        /// Access this computer from the network
        /// </summary>
        SeNetworkLogonRight,
        /// <summary>
        /// Act as part of the operating system
        /// </summary>
        SeTcbPrivilege,
        /// <summary>
        /// Add workstations to domain
        /// </summary>
        SeMachineAccountPrivilege,
        /// <summary>
        /// Adjust memory quotas for a process
        /// </summary>
        SeIncreaseQuotaPrivilege,
        /// <summary>
        /// Allow log on locally
        /// </summary>
        SeInteractiveLogonRight,
        /// <summary>
        /// Allow log on through Remote Desktop Services
        /// </summary>
        SeRemoteInteractiveLogonRight,
        /// <summary>
        /// Back up files and directories
        /// </summary>
        SeBackupPrivilege,
        /// <summary>
        /// Bypass traverse checking
        /// </summary>
        SeChangeNotifyPrivilege,
        /// <summary>
        /// Change the system time
        /// </summary>
        SeSystemtimePrivilege,
        /// <summary>
        /// Change the time zone
        /// </summary>
        SeTimeZonePrivilege,
        /// <summary>
        /// Create a pagefile
        /// </summary>
        SeCreatePagefilePrivilege,
        /// <summary>
        /// Create a token object
        /// </summary>
        SeCreateTokenPrivilege,
        /// <summary>
        /// Create global objects
        /// </summary>
        SeCreateGlobalPrivilege,
        /// <summary>
        /// Create permanent shared objects
        /// </summary>
        SeCreatePermanentPrivilege,
        /// <summary>
        /// Create symbolic links
        /// </summary>
        SeCreateSymbolicLinkPrivilege,
        /// <summary>
        /// Debug programs
        /// </summary>
        SeDebugPrivilege,
        /// <summary>
        /// Deny access this computer from the network
        /// </summary>
        SeDenyNetworkLogonRight,
        /// <summary>
        /// Deny log on as a batch job
        /// </summary>
        SeDenyBatchLogonRight,
        /// <summary>
        /// Deny log on as a service
        /// </summary>
        SeDenyServiceLogonRight,
        /// <summary>
        /// Deny log on locally
        /// </summary>
        SeDenyInteractiveLogonRight,
        /// <summary>
        /// Deny log on through Remote Desktop Services
        /// </summary>
        SeDenyRemoteInteractiveLogonRight,
        /// <summary>
        /// Enable computer and user accounts to be trusted for delegation
        /// </summary>
        SeEnableDelegationPrivilege,
        /// <summary>
        /// Force shutdown from a remote system
        /// </summary>
        SeRemoteShutdownPrivilege,
        /// <summary>
        /// Generate security audits
        /// </summary>
        SeAuditPrivilege,
        /// <summary>
        /// Impersonate a client after authentication
        /// </summary>
        SeImpersonatePrivilege,
        /// <summary>
        /// Increase a process working set
        /// </summary>
        SeIncreaseWorkingSetPrivilege,
        /// <summary>
        /// Increase scheduling priority
        /// </summary>
        SeIncreaseBasePriorityPrivilege,
        /// <summary>
        /// Load and unload device drivers
        /// </summary>
        SeLoadDriverPrivilege,
        /// <summary>
        /// Lock pages in memory
        /// </summary>
        SeLockMemoryPrivilege,
        /// <summary>
        /// Log on as a batch job
        /// </summary>
        SeBatchLogonRight,
        /// <summary>
        /// Log on as a service
        /// </summary>
        SeServiceLogonRight,
        /// <summary>
        /// Manage auditing and security log
        /// </summary>
        SeSecurityPrivilege,
        /// <summary>
        /// Modify an object label
        /// </summary>
        SeRelabelPrivilege,
        /// <summary>
        /// Modify firmware environment values
        /// </summary>
        SeSystemEnvironmentPrivilege,
        /// <summary>
        /// Perform volume maintenance tasks
        /// </summary>
        SeManageVolumePrivilege,
        /// <summary>
        /// Profile single process
        /// </summary>
        SeProfileSingleProcessPrivilege,
        /// <summary>
        /// Profile system performance
        /// </summary>
        SeSystemProfilePrivilege,
        /// <summary>
        /// "Read unsolicited input from a terminal device"
        /// </summary>
        SeUnsolicitedInputPrivilege,
        /// <summary>
        /// Remove computer from docking station
        /// </summary>
        SeUndockPrivilege,
        /// <summary>
        /// Replace a process level token
        /// </summary>
        SeAssignPrimaryTokenPrivilege,
        /// <summary>
        /// Restore files and directories
        /// </summary>
        SeRestorePrivilege,
        /// <summary>
        /// Shut down the system
        /// </summary>
        SeShutdownPrivilege,
        /// <summary>
        /// Synchronize directory service data
        /// </summary>
        SeSyncAgentPrivilege,
        /// <summary>
        /// Take ownership of files or other objects
        /// </summary>
        SeTakeOwnershipPrivilege,
        /// <summary>
        /// obtain an impersonation token for another user in the same session
        /// </summary>
        SeDelegateSessionUserImpersonatePrivilege
    }

    /// <summary>
    /// Type of delegation that an account can be configured with
    /// </summary>
    [Flags]
    public enum DelegationType : uint
    {
        /// <summary>
        /// Is not applicable, configurable or available for this account
        /// </summary>
        None = 0,
        /// <summary>
        /// Kerberos delegation, to only those Service Principal Names detailed in the DelegationEndpoints
        /// </summary>
        ConstrainedKerberos = 1,
        /// <summary>
        /// Kerberos delegation, to any known Service Principal Name in the trusted Domain
        /// </summary>
        KerberosToAny = 2,
        /// <summary>
        /// Any protocol, not just kerberos, can attempt delegation
        /// </summary>
        ConstrainedUsingAnyProtocol = 4
    }
    /// <summary>
    /// Indicates the type of account that an SID represents
    /// </summary>
    public enum SID_NAME_USE
    {
        /// <summary>
        /// Unknown
        /// </summary>
        SidTypeUnknown = 0,
        /// <summary>
        /// Local User account
        /// </summary>
        SidTypeUser = 1,
        /// <summary>
        /// Group account
        /// </summary>
        SidTypeGroup,
        /// <summary>
        /// Domain account
        /// </summary>
        SidTypeDomain,
        /// <summary>
        /// Alias account
        /// </summary>
        SidTypeAlias,
        /// <summary>
        /// Built-In windows group
        /// </summary>
        SidTypeWellKnownGroup,
        /// <summary>
        /// Deleted account
        /// </summary>
        SidTypeDeletedAccount,
        /// <summary>
        /// Invalid type
        /// </summary>
        SidTypeInvalid,
        /// <summary>
        /// Computer Account
        /// </summary>
        SidTypeComputer
    }
}
