//using System.ComponentModel;

namespace WinDevOps
{
    /// <summary> 
    /// Sets the default level of impersonation for applications that do not call CoInitializeSecurity.
    /// </summary>
    /// <remarks>
    /// It is not recommended that you change this value, because this will affect all COM server applications that do not set their own
    /// process-wide security, and might prevent them from working properly. If you are changing this value to affect the security settings
    /// for a particular COM application, then you should instead change the process-wide security settings for that particular COM application.
    /// </remarks>
    public enum EnumLegacyImpersonationLevel
    {
        /// <summary>
        /// The client is anonymous to the server. The server can impersonate the client, but the impersonation token (a local credential)
        /// does not contain any information about the client
        /// </summary>
        //[Description("Anonymous")]
        RPC_C_IMP_LEVEL_ANONYMOUS = 1,
        /// <summary>
        /// The server can obtain the client's identity and can impersonate the client to do ACL checks. Default
        /// </summary>
        //[Description("Identify")]
        RPC_C_IMP_LEVEL_IDENTIFY = 2,
        /// <summary>
        /// The server can impersonate the client while acting on its behalf, although with restrictions.
        /// The server can access resources on the same computer as the client.
        /// If the server is on the same computer as the client, it can access network resources as the client.
        /// If the server is on a computer different from the client, it can access only resources that are on the same computer as the server.
        /// This is the default setting for COM+ server applications.
        /// </summary>
        //[Description("Impersonate")]
        RPC_C_IMP_LEVEL_IMPERSONATE = 3,
        /// <summary>
        /// The server can impersonate the client while acting on its behalf, whether on the same computer as the client.
        /// During impersonation, the client's credentials (both those with local and those with network validity) can be passed to any number of machines.
        /// </summary>
        //[Description("Delegate")]
        RPC_C_IMP_LEVEL_DELEGATE = 4
    }
}
