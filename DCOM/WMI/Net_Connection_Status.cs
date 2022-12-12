namespace WinDevOps.WMI
{
    /// <summary>
    /// Connection state of a network card
    /// </summary>
    /// <remarks>
    /// <see href="https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-networkadapter"/></remarks>
    public enum Net_Connection_Status
    {
        Disconnected = 0,
        Connecting = 1,
        Connected = 2,
        Disconnecting = 3,
        Hardware_not_present = 4,
        Hardware_disabled = 5,
        Hardware_malfunction = 6,
        Media_disconnected = 7,
        Authenticating = 8,
        Authentication_succeeded = 9,
        Authentication_failed = 10,
        Invalid_address = 11,
        Credentials_required = 12
    }
}