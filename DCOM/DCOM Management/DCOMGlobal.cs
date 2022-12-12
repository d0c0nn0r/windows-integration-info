using System;
using System.Collections.Generic;
using System.Security.AccessControl;
using COMAdmin;
using DCOM.Global;
using Microsoft.Win32;

namespace WinDevOps
{
    /// <summary>
    /// This object holds computer-level DCOM settings information. 
    /// </summary>
    public class DCOMGlobal : IEquatable<DCOMGlobal>, IComparable<DCOMGlobal>
    {
        private readonly COMAdminCatalog _comAdmin;
        private COMAdminCatalogCollection _comCollection;
        /// <summary>
        /// store for ComputerName
        /// </summary>
        private readonly string _target;
        /// <summary>
        /// Local computer COMAdminCatalogObject item
        /// </summary>
        private COMAdminCatalogObject _localComp;

        /// <summary>
        /// Automatically publish changes to system. If true, using property setters will commit change directly to system.
        /// If false, user must use <see cref="Commit"/> to publish their changes manually.
        /// Default= false
        /// </summary>
        /// <remarks>
        /// When set to true, any changes at that time are then automatically pushed to the system.
        /// </remarks>
        public bool AutoCommit
        {
            get => _autoCommit;
            set
            {
                _autoCommit = value;
                if(value) Commit(); //push changes, when set to true
            }
        }

        /// <summary>
        /// Remote server name used by application proxies by default.
        /// </summary>
        public string ApplicationProxyRSN
        {
            get
            {
                return (string)_localComp.Value["ApplicationProxyRSN"];
            }

            set
            {
                _localComp.Value["ApplicationProxyRSN"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }

        /// <summary>
        /// A description of the computer.
        /// </summary>
        public string Description
        {
            get { return _localComp.Value["Description"].ToString(); }
            set {
                _localComp.Value["Description"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// The name of the computer. Extra spaces at the beginning and end of the string are stripped out. This property is returned when the Key or Name property method is called on an object of this collection.
        /// </summary>
        public string ComputerName => _localComp.Value["Name"].ToString();

        /// <summary>
        /// Set to True to enable DCOM on the computer.
        /// </summary>
        public bool DCOMEnabled
        {
            get => (bool)_localComp.Value["DCOMEnabled"];
            set {
                _localComp.Value["DCOMEnabled"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// Indicates whether COM Internet Services is enabled.
        /// </summary>
        public bool CISEnabled
        {
            get => (bool)_localComp.Value["CISEnabled"];
            set {
                _localComp.Value["CISEnabled"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }

        /// <summary>
        /// Authentication level used by applications that have Authentication set to Default. Values correspond to the Remote Procedure Call (RPC) authentication settings.
        /// </summary>
        public EnumLegacyAuthenticationLevel DefaultAuthenticationLevel
        {
            get
            {
                int i = (int)_localComp.Value["DefaultAuthenticationLevel"];
                return (EnumLegacyAuthenticationLevel) i;
            } 
            set {
                _localComp.Value["DefaultAuthenticationLevel"] = (int)value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// Impersonation level to allow if one is not set.
        /// </summary>
        public EnumLegacyImpersonationLevel DefaultImpersonationLevel
        {
            get
            {
                int i = (int)_localComp.Value["DefaultImpersonationLevel"];
                return (EnumLegacyImpersonationLevel)i;
            }
            set {
                _localComp.Value["DefaultImpersonationLevel"] = (int)value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// Determines whether the default type of port provided should be Internet (True) or intranet (False).
        /// </summary>
        public bool DefaultToInternetPorts
        {
            get => (bool)_localComp.Value["DefaultToInternetPorts"];
            set {
                _localComp.Value["DefaultToInternetPorts"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }

        /// <summary>
        /// Indicates whether the user of the partition mappings is checked into the domain store.
        /// </summary>
        public bool DSPartitionLookupEnabled
        {
            get => (bool)_localComp.Value["DSPartitionLookupEnabled"];
            set {
                _localComp.Value["DSPartitionLookupEnabled"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// Determines whether the ports listed in the Ports property are to be used for Internet (True) or for intranet (False).
        /// </summary>
        public bool InternetPortsListed
        {
            get => (bool)_localComp.Value["InternetPortsListed"];
            set {
                _localComp.Value["InternetPortsListed"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// Set to True if the computer is a router for the component load balancing (CLB) service.
        /// This property can be set to True only if the component load balancing service is currently installed on the computer; otherwise,
        /// it errors with COMADMIN_E_REQUIRES_DIFFERENT_PLATFORM.
        /// </summary>
        public bool IsRouter
        {
            get => (bool)_localComp.Value["IsRouter"];
            set {
                _localComp.Value["IsRouter"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        
        /// <summary>
        /// The CLSID of the object to balance.
        /// </summary>
        public string LoadBalancingCLSID
        {
            get => (string)_localComp.Value["LoadBalancingCLSID"];
            set {
                _localComp.Value["LoadBalancingCLSID"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// Indicates whether the user of the partition mappings is checked into the local store.
        /// </summary>
        public bool LocalPartitionLookupEnabled
        {
            get => (bool)_localComp.Value["LocalPartitionLookupEnabled"];
            set {
                _localComp.Value["LocalPartitionLookupEnabled"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// Indicates whether COM+ partitions can be used on the local computer. If this property is False, any attempt to use COM+ partitions results in an error.
        /// </summary>
        public bool PartitionsEnabled
        {
            get => (bool)_localComp.Value["PartitionsEnabled"];
            set {
                _localComp.Value["PartitionsEnabled"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// A string describing ports that are for either Internet or intranet use, depending on the <see cref="InternetPortsListed"/> property; for example, "500-599: 600-800".
        /// </summary>
        public string Ports
        {
            get => (string)_localComp.Value["Ports"];
            set {
                _localComp.Value["Ports"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// Enables use of resource dispensers.
        /// </summary>
        public bool ResourcePoolingEnabled
        {
            get => (bool)_localComp.Value["ResourcePoolingEnabled"];
            set {
                _localComp.Value["ResourcePoolingEnabled"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// Controls whether the RPC IIS proxy is enabled.
        /// The RPC IIS proxy is used in conjunction with IIS to forward calls to the RPC mechanism from IIS and is one of the core pieces of
        /// COM Internet Services, which is enabled by setting CISEnabled to True. For more information on RPCProxyEnabled, see <see href="https://docs.microsoft.com/en-us/windows/desktop/Rpc/rpc-over-http-security">HTTP RPC Security</see>.
        /// </summary>
        public bool RPCProxyEnabled
        {
            get
            {
                try
                {
                    var x = _localComp.Value["RPCProxyEnabled"];
                    if (x == null) return false; // default
                    return (bool) x;
                }
                catch(ArgumentException)
                {
                    return false;
                }
            }
            set {
                _localComp.Value["RPCProxyEnabled"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }

        /// <summary>
        /// Enforces in DCOM computers that cross-process calls to IUnknown::AddRef and IUnknown::Release methods are secured.
        /// </summary>
        public bool SecureReferencesEnabled
        {
            get => (bool)_localComp.Value["SecureReferencesEnabled"];
            set {
                _localComp.Value["SecureReferencesEnabled"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// Set to True if security tracking is enabled on objects.
        /// </summary>
        public bool SecurityTrackingEnabled
        {
            get => (bool)_localComp.Value["SecurityTrackingEnabled"];
            set {
                _localComp.Value["SecurityTrackingEnabled"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// Determines how the software restriction policy (SRP) handles activate-as-activator connections.
        /// If set to True, the SRP trust level that is configured for the server object is compared with the SRP trust level of the client
        /// object and the higher (more stringent) trust level is used to run the server object. If set to False, the server object runs with the
        /// SRP trust level of the client object, regardless of the SRP trust level with which the server is configured.
        /// </summary>
        public bool SRPActivateAsActivatorChecks
        {
            get => (bool)_localComp.Value["SRPActivateAsActivatorChecks"];
            set {
                _localComp.Value["SRPActivateAsActivatorChecks"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// Determines how the software restriction policy (SRP) handles attempted connections to existing processes.
        /// If set to False, attempts to connect to running objects are not checked for appropriate SRP trust levels.
        /// If set to True, the running object must have an equal or higher (more stringent) SRP trust level than the client object.
        /// For example, a client object with an Unrestricted SRP trust level cannot connect to a running object with a Disallowed SRP trust level.
        /// </summary>
        public bool SRPRunningObjectChecks
        {
            get => (bool)_localComp.Value["SRPRunningObjectChecks"];
            set {
                _localComp.Value["SRPRunningObjectChecks"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }
        /// <summary>
        /// Should be set to a sufficient value in seconds if you are doing numerous operations within a transaction.
        /// The default time-out period is 60 seconds, and the maximum time-out period is 3600 seconds (1 hour).
        /// Setting this property to 0 disables transaction time-outs.
        /// This property can be overridden by individual components by using the ComponentTransactionTimeout property of the Components collection.
        /// </summary>
        public int TransactionTimeout
        {
            get => (int)_localComp.Value["TransactionTimeout"];
            set
            {
                _localComp.Value["TransactionTimeout"] = value;
                if (AutoCommit)
                {
                    Commit();
                }
            }
        }

        /// <inheritdoc cref="DCOMApplication.GetDcomApplications"/>
        public List<DCOMApplication> GetApplications(string appName = null)
        {
            return DCOMApplication.GetDcomApplications(ComputerName, appName, this);
        }

        private List<DCOMProtocol> _protocols;
        private bool _autoCommit;

        /// <summary>
        /// Contains a list of the protocols to be used by DCOM. It contains an object for each protocol.
        /// </summary>
        public List<DCOMProtocol> DcomProtocols => _protocols;

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="computerName">Network name of remote computer to read COM settings from. Leave blank to load from local computer</param>
        public DCOMGlobal(string computerName = null)
        {
            _comAdmin = new COMAdminCatalog();
            _target = computerName;
            Refresh();
        }
        /// <summary>
        /// Reload COM settings
        /// </summary>
        public void Refresh()
        {
            _comAdmin.Connect(!string.IsNullOrEmpty(_target) ? _target : "localhost");
            _comCollection = (COMAdminCatalogCollection)_comAdmin.GetCollection("LocalComputer");

            _comCollection.Populate();
            _localComp = (COMAdminCatalogObject)_comCollection.Item[0];

            COMAdminCatalogCollection proto = (COMAdminCatalogCollection)_comAdmin.GetCollection("DCOMProtocols");
            proto.Populate();
            _protocols = new List<DCOMProtocol>(proto.Count);
            foreach (COMAdminCatalogObject o in proto)
            {
                _protocols.Add(new DCOMProtocol(o));
            }

            RefreshSecurity();
        }

        /// <summary>
        /// Commit DCOM changes to system.
        /// <see cref="AutoCommit"/>
        /// </summary>
        public void Commit()
        {
            _comCollection.SaveChanges();
        }


        /// <inheritdoc />
        public bool Equals(DCOMGlobal other)
        {
            StringEqualNullEmptyComparer sC = new StringEqualNullEmptyComparer(); //comparer to handle string, null, empties, foreign language
            var r = sC.Equals(ComputerName, other.ComputerName, StringComparison.OrdinalIgnoreCase) &&
                   sC.Equals(Description, other.Description, StringComparison.OrdinalIgnoreCase) &&
                   sC.Equals(ApplicationProxyRSN, other.ApplicationProxyRSN, StringComparison.OrdinalIgnoreCase) &&
                   sC.Equals(LoadBalancingCLSID, other.LoadBalancingCLSID, StringComparison.OrdinalIgnoreCase) &&
                   sC.Equals(Ports, other.Ports, StringComparison.OrdinalIgnoreCase) &&
                   DefaultToInternetPorts == other.DefaultToInternetPorts &&
                   InternetPortsListed == other.DefaultToInternetPorts &&
                   CISEnabled == other.CISEnabled &&
                   DCOMEnabled == other.DCOMEnabled &&
                   DSPartitionLookupEnabled == other.DSPartitionLookupEnabled &&
                   IsRouter == other.IsRouter &&
                   LocalPartitionLookupEnabled == other.LocalPartitionLookupEnabled &&
                   PartitionsEnabled == other.PartitionsEnabled &&
                   RPCProxyEnabled == other.RPCProxyEnabled &&
                   ResourcePoolingEnabled == other.ResourcePoolingEnabled &&
                   SRPActivateAsActivatorChecks == other.SRPActivateAsActivatorChecks &&
                   SRPRunningObjectChecks == other.SRPRunningObjectChecks &&
                   SecureReferencesEnabled == other.SecureReferencesEnabled &&
                   SecurityTrackingEnabled == other.SecurityTrackingEnabled &&
                   TransactionTimeout == other.TransactionTimeout;
            if (!r) return false;
            return ProtocolListEquality(other);
        }

        private bool SimpleEquality(DCOMGlobal other)
        {
            StringEqualNullEmptyComparer sC = new StringEqualNullEmptyComparer(); //comparer to handle string, null, empties, foreign language
            return sC.Equals(ApplicationProxyRSN, other.ApplicationProxyRSN, StringComparison.OrdinalIgnoreCase) &&
                   sC.Equals(LoadBalancingCLSID, other.LoadBalancingCLSID, StringComparison.OrdinalIgnoreCase) &&
                   sC.Equals(Ports, other.Ports, StringComparison.OrdinalIgnoreCase) &&
                   DefaultToInternetPorts == other.DefaultToInternetPorts &&
                   InternetPortsListed == other.DefaultToInternetPorts &&
                   CISEnabled == other.CISEnabled &&
                   DCOMEnabled == other.DCOMEnabled &&
                   DSPartitionLookupEnabled == other.DSPartitionLookupEnabled &&
                   IsRouter == other.IsRouter &&
                   LocalPartitionLookupEnabled == other.LocalPartitionLookupEnabled &&
                   PartitionsEnabled == other.PartitionsEnabled &&
                   RPCProxyEnabled == other.RPCProxyEnabled &&
                   ResourcePoolingEnabled == other.ResourcePoolingEnabled &&
                   SRPActivateAsActivatorChecks == other.SRPActivateAsActivatorChecks &&
                   SRPRunningObjectChecks == other.SRPRunningObjectChecks &&
                   SecureReferencesEnabled == other.SecureReferencesEnabled &&
                   SecurityTrackingEnabled == other.SecurityTrackingEnabled &&
                   TransactionTimeout == other.TransactionTimeout;
        }
        /// <summary>
        /// Compare lists
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        private bool ProtocolListEquality(DCOMGlobal other)
        {
            var cnt = new Dictionary<DCOMProtocol, int>();
            foreach (DCOMProtocol s in DcomProtocols)
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
            foreach (DCOMProtocol s in other.DcomProtocols)
            {
                if (cnt.ContainsKey(s))
                {
                    cnt[s]--;
                }
                else
                {
                    return false; //found mismatch. Short-circuit out 'false' result
                }
            }

            foreach (var val in cnt.Values)
            {
                if (val != 0) return false;
            }
            //return cnt.Values.All(c => c == 0); //all processed. See if any from source list that did not get matched
            return true;
        }

        /// <inheritdoc />
        public int CompareTo(DCOMGlobal other)
        {
            return String.Compare(ComputerName, other.ComputerName, StringComparison.Ordinal);
        }

        /// <inheritdoc />
        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 13;
                hash = (hash * 7) ^ ComputerName.GetHashCode();
                return hash;
            }
        }

        /// <summary>
        /// Using the <paramref name="other"></paramref> dcom settings, replicate all those settings to this object and save them to the machine
        /// </summary>
        /// <param name="other">Target <see cref="DCOMGlobal"/> to copy settings from</param>
        /// <param name="overwriteACL">If true, then <see cref="DefaultLaunchAndActivation"/>, <see cref="DefaultAccessPermissions"/>, <see cref="LimitsLaunchAndActivation"/>, <see cref="LimitsAccessPermissions"/>
        /// are overwritten completely. Any existing rules not in the target will be removed.
        /// If false, only additional rules are added and privileges for matching accounts are replaced.</param>
        /// <returns></returns>
        /// <remarks>
        /// Copying of <see cref="DcomProtocols"/> is not supported by this copy function.
        ///
        /// When copying Access Control Lists, the <see cref="CopyAccessControlList"/> method is called for
        /// each, Launch And Activation and Access Permissions.
        /// Each sub-category (<see cref="DcomPermissionOption.Default"/>, and <see cref="DcomPermissionOption.Limits"/>) is also updated.
        /// By default, <paramref name="overwriteACL"/> is <code>false</code>, so security is merged without removing any pre-existing accounts from the ACL.
        /// If <paramref name="overwriteACL"/> is <code>true</code>, then ACL for each is completely replaced with that defined in <paramref name="other"/>.
        /// </remarks>
        public bool CopyFrom(DCOMGlobal other, bool overwriteACL = false)
        {
            var preVal = AutoCommit;
            AutoCommit = false;
            ApplicationProxyRSN = other.ApplicationProxyRSN;
            CISEnabled = other.CISEnabled;
            DCOMEnabled = other.DCOMEnabled;
            DefaultAuthenticationLevel = other.DefaultAuthenticationLevel;
            DefaultImpersonationLevel = other.DefaultImpersonationLevel;
            InternetPortsListed = other.InternetPortsListed;
            IsRouter = other.IsRouter;
            LoadBalancingCLSID = other.LoadBalancingCLSID;
            LocalPartitionLookupEnabled = other.LocalPartitionLookupEnabled;
            PartitionsEnabled = other.PartitionsEnabled;
            Ports = other.Ports;
            RPCProxyEnabled = other.RPCProxyEnabled;
            ResourcePoolingEnabled = other.ResourcePoolingEnabled;
            SRPActivateAsActivatorChecks = other.SRPActivateAsActivatorChecks;
            SRPRunningObjectChecks = other.SRPRunningObjectChecks;
            SecureReferencesEnabled = other.SecureReferencesEnabled;
            SecurityTrackingEnabled = other.SecurityTrackingEnabled;
            TransactionTimeout = other.TransactionTimeout;
            DSPartitionLookupEnabled = other.DSPartitionLookupEnabled;
            DefaultToInternetPorts = other.DefaultToInternetPorts;
            Commit();
            CopyAccessControlList(other, overwriteACL);
            CopyAccessControlList(other, overwriteACL, DcomSecurityTypes.Access, DcomPermissionOption.Limits);
            CopyAccessControlList(other, overwriteACL, DcomSecurityTypes.Launch);
            CopyAccessControlList(other, overwriteACL, DcomSecurityTypes.Launch, DcomPermissionOption.Limits);
            AutoCommit = preVal;
            return SimpleEquality(other);

        }

        /// <summary>
        /// Access Permissions Default Access Control List
        /// </summary>
        public List<DCOMAce> DefaultAccessPermissions { get; private set; }
        /// <summary>
        /// Access Permissions Limits Access Control List
        /// </summary>
        public List<DCOMAce> LimitsAccessPermissions { get; private set; }

        /// <summary>
        /// Launch And Activation Permissions Default Access Control List
        /// </summary>
        public List<DCOMAce> DefaultLaunchAndActivation { get; private set; }
        /// <summary>
        /// Launch And Activation Permissions Default Access Control List
        /// </summary>
        public List<DCOMAce> LimitsLaunchAndActivation { get; private set; }

        /// <summary>
        /// Retrieve DCOM security settings from the target computer registry
        /// </summary>
        /// <returns></returns>
        private void RefreshSecurity()
        {
            var _verboseMessages = new List<string>();
            
            RegistryKey rootKey = string.IsNullOrEmpty(ComputerName) ? Registry.LocalMachine :
                RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, ComputerName);
            
            using (RegistryKey registryKey = rootKey.OpenSubKey("SOFTWARE")?.OpenSubKey("Microsoft")?.OpenSubKey("OLE"))
            {
                if (registryKey == null)
                {
                    throw new ArgumentNullException(
                        "The registry key 'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\OLE' was not found");
                }
                _verboseMessages.Add($"{Environment.MachineName} : Attempting to read DCOM registry keys");
                byte[] defLaunchPermission = (byte[])registryKey.GetValue("DefaultLaunchPermission");
                byte[] defAccessPermis = (byte[])registryKey.GetValue("DefaultAccessPermission");
                byte[] limLaunchPermission = (byte[])registryKey.GetValue("MachineLaunchRestriction");
                byte[] limAccessPermis = (byte[])registryKey.GetValue("MachineAccessRestriction");
                RawSecurityDescriptor newDescriptor;
                if (limAccessPermis == null || limAccessPermis.Length == 0)
                {
                    throw new AccessViolationException("Unable to retrieve Registry Settings");
                }
                if (limLaunchPermission == null || limLaunchPermission.Length == 0)
                {
                    throw new AccessViolationException("Unable to retrieve Registry Settings");
                }

                //Access Permissions Default
                _verboseMessages.Add(defAccessPermis != null ? $"{Environment.MachineName} : processing Access Permissions Default ACL" : $"{Environment.MachineName} : Access Permissions Default not found. Assuming machine defaults....");
                newDescriptor = defAccessPermis != null ? new RawSecurityDescriptor(defAccessPermis, 0) : new RawSecurityDescriptor(
                    "O:BAG:BAD:(A;;CCDCLC;;;PS)(A;;CCDC;;;SY)(A;;CCDCLC;;;BA)");
                DefaultAccessPermissions = DcomUtils.DeserializeACLtoACE(DcomSecurityTypes.Access, newDescriptor,ComputerName, DcomPermissionOption.Default);
                
                //Access Permissions Limits
                _verboseMessages.Add($"{Environment.MachineName} : processing Access Permissions Limits ACL");
                newDescriptor = new RawSecurityDescriptor(limAccessPermis, 0);
                LimitsAccessPermissions = DcomUtils.DeserializeACLtoACE(DcomSecurityTypes.Access, newDescriptor, ComputerName, DcomPermissionOption.Limits);
                

                //LaunchAndActivation Permissions Default
                _verboseMessages.Add(defLaunchPermission != null ? $"{Environment.MachineName} : processing LaunchAndActivation Permissions Default ACL" : $"{Environment.MachineName} : LaunchAndActivation Permissions Default not found. Assuming machine defaults....");
                newDescriptor = defLaunchPermission != null ? new RawSecurityDescriptor(defLaunchPermission, 0) : new RawSecurityDescriptor(
                    "O:BAG:BAD:(A;;CCDCLC;;;PS)(A;;CCDC;;;SY)(A;;CCDCLC;;;BA)");
                DefaultLaunchAndActivation = DcomUtils.DeserializeACLtoACE(DcomSecurityTypes.Launch, newDescriptor, ComputerName, DcomPermissionOption.Default);
                
                //LaunchAndActivation Permissions Limits
                _verboseMessages.Add($"{Environment.MachineName} : processing LaunchAndActivation Permissions Limits ACL");
                newDescriptor = new RawSecurityDescriptor(limLaunchPermission, 0);
                LimitsLaunchAndActivation = DcomUtils.DeserializeACLtoACE(DcomSecurityTypes.Launch, newDescriptor, ComputerName, DcomPermissionOption.Limits);
                
            }
        
        }


        /// <summary>
        /// Update the <see cref="DefaultAccessPermissions"/>, <see cref="DefaultLaunchAndActivation"/>, <see cref="LimitsLaunchAndActivation"/> or <see cref="LimitsAccessPermissions"/>
        /// of this computer, by copying the ACL from <paramref name="copyFrom"/>
        /// defined on the other application.
        /// To replace ACL completely, replacing/removing any of the current settings, set <paramref name="overwrite"/> to true.
        /// </summary>
        /// <param name="copyFrom">The <see cref="DCOMGlobal"/> to copy Privileges from</param>
        /// <param name="overwrite">Set to true, to replace and remove privileges defined locally, entirely with those in <paramref name="copyFrom"/>.
        /// If false, missing accounts are added, and privileges for matching accounts are updated to match <paramref name="copyFrom"/></param>
        /// <param name="type">The Access Control List to update. Can be either Launch or Access. Use <paramref name="category"/> to define whether the LIMITS or DEFAULT ACL is to be updated.</param>
        /// <param name="category">Select whether comparing LIMITS or DEFAULT access control lists</param>
        /// <exception cref="NotSupportedException">Thrown if <paramref name="type"/> value provided is Config. This is not supported for ACE copying</exception>
        /// <returns></returns>
        /// <exception cref="AggregateException">Exceptions representing ACE rules that failed to be copied successfully.</exception>
        public bool CopyAccessControlList(DCOMGlobal copyFrom, bool overwrite = false, DcomSecurityTypes type = DcomSecurityTypes.Access, DcomPermissionOption category = DcomPermissionOption.Default)
        {
            if (type == DcomSecurityTypes.Config)
            {
                throw new NotSupportedException(
                    "DcomSecurityTypes.Config is not supported for copying Access Control Lists");
            }

            var retList = new List<bool>();
            
            var toRemove = type == DcomSecurityTypes.Access ? GetMismatchedAccessPermissions(copyFrom, category) : GetMismatchedLaunchPermissions(copyFrom, category);
            var toAdd = type == DcomSecurityTypes.Access ? copyFrom.GetMismatchedAccessPermissions(this, category) : copyFrom.GetMismatchedLaunchPermissions(this, category);
            foreach (var dcomAce in toAdd)
            {

                var r = DcomUtils.AddComputerRights(dcomAce.SecurityType, dcomAce.SID,
                    dcomAce.ConvertToMachineRightsList()?.ToArray(), category, dcomAce.AccessType, ComputerName);
                if (!r) throw new Exception($"Failed to apply Access Control List, Type: {dcomAce.AccessType}, Category: {dcomAce.Category}, SecurityType: {dcomAce.SecurityType}");
                retList.Add(r);
            }

            if (overwrite)
            {
                foreach (var dcomAce in toRemove)
                {
                    var r = DcomUtils.RemoveComputerRights(dcomAce.SecurityType, dcomAce.SID, category, ComputerName);
                    if (!r) throw new Exception($"Failed to apply Access Control List, Type: {dcomAce.AccessType}, Category: {dcomAce.Category}, SecurityType: {dcomAce.SecurityType}");
                    retList.Add(r);
                }
            }
            
            RefreshSecurity();
            return retList.TrueForAll(a => a);
        }

        /// <summary>
        /// Compare <seealso cref="DefaultLaunchAndActivation"/> or <seealso cref="LimitsLaunchAndActivation"/> between two different <see cref="DCOMGlobal"/>, and see it they're the same.
        /// Equality is limited to
        /// </summary>
        /// <param name="other"></param>
        /// <param name="category">Select whether comparing LIMITS or DEFAULT access control lists</param>
        /// <returns></returns>
        public bool LaunchPermissionEquality(DCOMGlobal other, DcomPermissionOption category = DcomPermissionOption.Default)
        {
            List<DCOMAce> rules = null;
            List<DCOMAce> otherRules = null;
            if (category == DcomPermissionOption.None)
            {
                throw new NotSupportedException($"{nameof(category)} value provided not supported");
            }
            switch (category)
            {
                case DcomPermissionOption.Default:
                    rules = DefaultLaunchAndActivation;
                    otherRules = other.DefaultLaunchAndActivation;
                    break;
                case DcomPermissionOption.Limits:
                    rules = LimitsLaunchAndActivation;
                    otherRules = other.LimitsLaunchAndActivation;
                    break;
            }
            if (rules == null || rules.Count <= 0)
            {
                if (otherRules == null || otherRules.Count <=1)
                {
                    return true;
                }
                return false;
            }

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
            foreach (var s in otherRules)
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
        /// Compare <seealso cref="DefaultAccessPermissions"/> or <seealso cref="LimitsAccessPermissions"/> between two different <see cref="DCOMGlobal"/>, and see it they're the same
        /// </summary>s
        /// <param name="other"></param>
        /// <param name="category">Select whether comparing LIMITS or DEFAULT access control lists</param>
        /// <returns></returns>
        public bool AccessPermissionEquality(DCOMGlobal other, DcomPermissionOption category = DcomPermissionOption.Default)
        {
            List<DCOMAce> rules = null;
            List<DCOMAce> otherRules = null;
            if (category == DcomPermissionOption.None)
            {
                throw new NotSupportedException($"{nameof(category)} value provided not supported");
            }
            switch (category)
            {
                case DcomPermissionOption.Default:
                    rules = DefaultAccessPermissions;
                    otherRules = other.DefaultAccessPermissions;
                    break;
                case DcomPermissionOption.Limits:
                    rules = LimitsAccessPermissions;
                    otherRules = other.LimitsAccessPermissions;
                    break;
            }
            if (rules == null || rules.Count <=0)
            {
                if (otherRules == null || otherRules.Count <= 0)
                {
                    return true;
                }
                return false;
            }

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
            foreach (var s in otherRules)
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
        /// Get unique list of the <seealso cref="DefaultLaunchAndActivation"/> or <seealso cref="LimitsLaunchAndActivation"/> objects of that don't match the Launch Permissions of the <paramref name="other"/> object.
        /// </summary>
        /// <param name="other"></param>
        /// <param name="category">Select whether comparing LIMITS or DEFAULT access control lists</param>
        /// <returns></returns>
        /// <remarks>
        /// Return list does not contains any <see cref="DCOMAce"/> objects from the <paramref name="other"/> object
        /// </remarks>
        public List<DCOMAce> GetMismatchedLaunchPermissions(DCOMGlobal other, DcomPermissionOption category = DcomPermissionOption.Default)
        {
            return GetMismatched(other, DcomSecurityTypes.Launch, category);
        }

        /// <summary>
        /// Get unique list of the <seealso cref="DefaultAccessPermissions"/> or <seealso cref="LimitsAccessPermissions"/> objects of that don't match the Access Permissions of the <paramref name="other"/> object.
        /// </summary>
        /// <param name="other"></param>
        /// <param name="category">Select whether comparing LIMITS or DEFAULT access control lists</param>
        /// <returns></returns>
        /// <remarks>
        /// Return list does not contains any <see cref="DCOMAce"/> objects from the <paramref name="other"/> object
        /// </remarks>
        public List<DCOMAce> GetMismatchedAccessPermissions(DCOMGlobal other, DcomPermissionOption category = DcomPermissionOption.Default)
        {
            return GetMismatched(other, DcomSecurityTypes.Access, category);
        }

        /// <summary>
        /// Get Unique list of <see cref="DCOMAce"/> rules by comparing the Launch or Access permissions between objects
        /// </summary>
        /// <param name="other"></param>
        /// <param name="type"></param>
        /// <param name="category">Select whether comparing LIMITS or DEFAULT access control lists</param>
        /// <exception cref="NotSupportedException">Thrown if <paramref name="type"/> is set to CONFIG</exception>
        /// <returns></returns>
        /// <remarks>
        /// Return list does not contains any <see cref="DCOMAce"/> objects from the <paramref name="other"/> object
        /// </remarks>
        private List<DCOMAce> GetMismatched(DCOMGlobal other, DcomSecurityTypes type, DcomPermissionOption category = DcomPermissionOption.Default)
        {
            if (type == DcomSecurityTypes.Config)
            {
                throw new NotSupportedException(nameof(type));
            }

            if (category == DcomPermissionOption.None)
            {
                throw new NotSupportedException($"{nameof(category)} value provided not supported");
            }

            List<DCOMAce> rules = new List<DCOMAce>();
            List<DCOMAce> otherRules = new List<DCOMAce>();
            switch (type)
            {
                case DcomSecurityTypes.Access:
                    switch (category)
                    {
                        case DcomPermissionOption.Default:
                            rules = DefaultAccessPermissions;
                            otherRules = other.DefaultAccessPermissions;
                            break;
                        case DcomPermissionOption.Limits:
                            rules = LimitsAccessPermissions;
                            otherRules = other.LimitsAccessPermissions;
                            break;

                    }
                    break;
                case DcomSecurityTypes.Launch:
                    switch (category)
                    {
                        case DcomPermissionOption.Default:
                            if (DefaultLaunchAndActivation == null || DefaultLaunchAndActivation.Count <=0) return new List<DCOMAce>();
                            rules = DefaultLaunchAndActivation;
                            otherRules = other.DefaultLaunchAndActivation;
                            break;
                        case DcomPermissionOption.Limits:
                            if (LimitsLaunchAndActivation == null || LimitsLaunchAndActivation.Count <= 0) return new List<DCOMAce>();
                            rules = LimitsLaunchAndActivation;
                            otherRules = other.LimitsLaunchAndActivation;
                            break;
                    }
                    break;
            }

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
            foreach (var s in otherRules)
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
                if (key.Value == 1) lst.Add(key.Key);
            }

            return lst;
            //return cnt.Where(c => c.Value == 1).Select(c => c.Key).ToList(); //all processed. See if any from source list that did not get matched
        }


    }
}
