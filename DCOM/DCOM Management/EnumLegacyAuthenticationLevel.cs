
namespace WinDevOps
{
    /// <summary>
    /// Sets the default authentication level for applications that do not call CoInitializeSecurity.
    /// If this registry value is not present, the default authentication level established by the system is 2 (RPC_C_AUTHN_CONNECT).
    /// </summary>
    /// <remarks>
    /// It is not recommended that you change this value, because this will affect all COM server applications that do not set their own process-wide security, and might prevent them from working properly. If you are changing this value to affect the security settings for a particular COM application, then you should instead change the process-wide security settings for that particular COM application.
    /// </remarks> 
    public enum EnumLegacyAuthenticationLevel
    {
        /// <summary>
        /// Default is <see cref="RPC_C_AUTHN_LEVEL_CONNECT"/>
        /// </summary>
        //[Description("Default")]
        RPC_C_AUTHN_LEVEL_DEFAULT = 0,
        /// <summary>
        /// No authentication occurs
        /// </summary>
        //[Description("None")]
        RPC_C_AUTHN_LEVEL_NONE = 1,
        /// <summary>
        /// Authenticates credentials only when the connection is made
        /// </summary>
        //[Description("Connect")]
        RPC_C_AUTHN_LEVEL_CONNECT=2,
        /// <summary>
        /// Authenticates credentials at the beginning of every call
        /// </summary>
        //[Description("Call")]
        RPC_C_AUTHN_LEVEL_CALL = 3,
        /// <summary>
        /// Authenticates credentials and verifies that all call data is received. This is the default setting for COM+ server applications
        /// </summary>
        //[Description("Packet")]
        RPC_C_AUTHN_LEVEL_PKT = 4,
        /// <summary>
        /// Authenticates credentials and verifies that no call data has been modified in transit
        /// </summary>
        //[Description("Packet Integrity")]
        RPC_C_AUTHN_LEVEL_PKT_INTEGRITY = 5,
        /// <summary>
        /// Authenticates credentials and encrypts the packet, including the data and the sender's identity and signature
        /// </summary>
        //[Description("Packet Privacy")]
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY = 6
    }
}
