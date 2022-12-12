using System;
using System.Collections.Generic;
using System.Security.AccessControl;
using System.Security.Principal;
using DCOM.Global;
using Newtonsoft.Json;

namespace WinDevOps
{
    /// <summary>
    /// Access Control List / Permissions for DCOM computer-wide and application-specific settings
    /// </summary>
    public class DCOMAce : IEquatable<DCOMAce>
    {
        /// <summary>
        /// SID for the <see cref="User"/>
        /// </summary>
        [JsonIgnore]
        public readonly SecurityIdentifier SID;

        /// <summary>
        /// Account name that this Access Control List is for
        /// </summary>
        public string User { get; internal set; }
        
        /// <inheritdoc cref="System.Security.AccessControl.AccessControlType" />
        public AccessControlType AccessType { get; internal set; }

        /// <summary>
        /// Local Access permission. If true, then this permission is enabled. If false, then not enabled.
        /// If null, then this permission is not applicable.
        /// </summary>
        public bool? LocalAccess { get; internal set; }
        /// <summary>
        /// Remote Access permission. If true, then this permission is enabled. If false, then not enabled.
        /// If null, then this permission is not applicable.
        /// </summary>
        public bool? RemoteAccess { get; internal set; }
        /// <summary>
        /// Local Launch permission. If true, then this permission is enabled. If false, then not enabled.
        /// If null, then this permission is not applicable.
        /// </summary>
        public bool? LocalLaunch { get; internal set; }
        /// <summary>
        /// Remote Launch permission. If true, then this permission is enabled. If false, then not enabled.
        /// If null, then this permission is not applicable.
        /// </summary>
        public bool? RemoteLaunch { get; internal set; }
        /// <summary>
        /// Local Activation permission. If true, then this permission is enabled. If false, then not enabled.
        /// If null, then this permission is not applicable.
        /// </summary>
        public bool? LocalActivation { get; internal set; }
        /// <summary>
        /// Remote Activation permission. If true, then this permission is enabled. If false, then not enabled.
        /// If null, then this permission is not applicable.
        /// </summary>
        public bool? RemoteActivation { get; internal set; }
        /// <inheritdoc cref="DcomSecurityTypes"/>
        public readonly DcomSecurityTypes SecurityType;
        
        /// <inheritdoc cref="DcomPermissionOption"/>
        public readonly DcomPermissionOption Category;
        
        /// <summary>
        /// Constructor. <see cref="Category"/> will be set to NONE.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="sid"></param>
        internal DCOMAce(DcomSecurityTypes type, SecurityIdentifier sid)
        {
            SecurityType = type;
            Category = DcomPermissionOption.None;
            //_ace = ace;
            SID = sid;
        }

        /// <summary>
        /// Constructor, allows setting <see cref="Category"/>
        /// </summary>
        /// <param name="type"></param>
        /// <param name="category"></param>
        /// <param name="sid"></param>
        internal DCOMAce(DcomSecurityTypes type, DcomPermissionOption category, SecurityIdentifier sid)
        {
            SecurityType = type;
            Category = category;
            //_ace = ace;
            SID = sid;
        }

        /// <inheritdoc />
        public bool Equals(DCOMAce other)
        {
            StringEqualNullEmptyComparer comparer = new StringEqualNullEmptyComparer();
            return comparer.Equals(User, other.User, StringComparison.OrdinalIgnoreCase) &&
                   SecurityType.Equals(other.SecurityType) && Category.Equals(other.Category) &&
                   AccessType.Equals(other.AccessType) &&
                   LocalAccess.Equals(other.LocalAccess) && RemoteAccess.Equals(other.RemoteAccess) &&
                   LocalActivation.Equals(other.LocalActivation) &&
                   RemoteActivation.Equals(other.RemoteActivation) &&
                   LocalLaunch.Equals(other.LocalLaunch) && RemoteLaunch.Equals(other.RemoteLaunch);
        }

        /// <inheritdoc />
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;

            if (obj.GetType() != GetType()) return false;
            DCOMAce other = (DCOMAce) obj;
            return Equals(other);
        }

        /// <inheritdoc />
        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 13;
                hash = (hash * 7) ^ User.GetHashCode();
                hash = (hash * 7) ^ AccessType.GetHashCode();
                hash = (hash * 7) ^ SecurityType.GetHashCode(); 
                hash = (hash * 7) ^ Category.GetHashCode();
                return hash;
            }
        }
        /// <summary>
        /// Serialize this Access Control list as a <see cref="MachineDComAccessRights"/> array
        /// </summary>
        /// <returns></returns>
        public List<MachineDComAccessRights> ConvertToMachineRightsList()
        {
            List<MachineDComAccessRights> ret = new List<MachineDComAccessRights>();
            if (LocalActivation.HasValue && LocalActivation.Value)
            {
                if(!ret.Contains(MachineDComAccessRights.ActivateLocal))
                {
                    ret.Add(MachineDComAccessRights.ActivateLocal);
                }
            }
            if (RemoteActivation.HasValue && RemoteActivation.Value)
            {
                if (!ret.Contains(MachineDComAccessRights.ActivateRemote))
                {
                    ret.Add(MachineDComAccessRights.ActivateRemote);
                }
                
            }
            if ((LocalAccess.HasValue && LocalAccess.Value) || (LocalLaunch.HasValue && LocalLaunch.Value))
            {
                if (!ret.Contains(MachineDComAccessRights.Execute))
                {
                    ret.Add(MachineDComAccessRights.Execute);
                }
                if (!ret.Contains(MachineDComAccessRights.ExecuteLocal))
                {
                    ret.Add(MachineDComAccessRights.ExecuteLocal);
                }
            }
            if ((RemoteAccess.HasValue && RemoteAccess.Value) || (RemoteLaunch.HasValue && RemoteLaunch.Value))
            {
                if (!ret.Contains(MachineDComAccessRights.Execute))
                {
                    ret.Add(MachineDComAccessRights.Execute);
                }
                if (!ret.Contains(MachineDComAccessRights.ExecuteRemote))
                {
                    ret.Add(MachineDComAccessRights.ExecuteRemote);
                }
            }

            return ret;
        }

        /// <inheritdoc />
        public override string ToString()
        {
            return $"{User} : {string.Join("|", ConvertToMachineRightsList().ConvertAll(a => a.ToString()).ToArray())}";
        }
    }
}