using System;

namespace WinDevOps.WMI
{
    /// <summary>
    /// Modes supported by a network card.
    /// </summary>
    /// <remarks>
    /// <see href="https://docs.microsoft.com/en-us/previous-versions/windows/embedded/aa448106(v=msdn.10)?redirectedfrom=MSDN"/>
    /// </remarks>
    [Flags]
    public enum NDIS_PACKET_TYPE
    {
        NONE=0,
        DIRECTED = 1,
        MULTICAST = 2,
        ALL_MULTICAST = 4,
        BROADCAST = 8,
        SOURCE_ROUTING = 16,
        PROMISCUOUS = 32,
    }
}