namespace WinDevOps
{
    /// <summary>
    /// The user account to run the DCOM server on
    /// </summary>
    public enum DcomApplicationRunAs
    {
        /// <summary>
        /// Select "The launching user" if you want the DCOM server to run on the account of the launching user.
        /// </summary>
        LaunchingUser,
        /// <summary>
        /// Select "The interactive user" if you want the DCOM server to appear on the Windows desktop of the currently logged-on user
        /// </summary>
        InteractiveUser
    }
}