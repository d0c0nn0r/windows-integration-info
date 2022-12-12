using System.Collections.Generic;
using WinDevOps.WMI;

namespace WinDevOps.WMI
{
    /// <summary>
    /// Class for all hardware details of a computer system that is of importance
    /// </summary>
    public class ComputerInformation
    {
        public ComputerProperties ComputerSystem { get; internal set; }
        public OSProperties OperatingSystem { get; internal set; }

        public ChassisProperties Chassis { get; internal set; }

        public BIOSProperties Bios { get; internal set; }
        public List<MemoryArrayProperties> MemoryArrayProperties { get; internal set; }
        public List<MemoryProperties> MemoryProperties { get; internal set; }
        public ProcessorProperties ProcessorProperties { get; internal set; }

        public List<DiskProperties> Disks { get; internal set; }
        public List<NetworkCardProperties> NetworkCards { get; internal set; }

        public TimeProperties Time { get; internal set; }
        public ComputerInformation()
        {
            Disks = new List<DiskProperties>();
            NetworkCards = new List<NetworkCardProperties>();
            MemoryProperties = new List<MemoryProperties>();
            MemoryArrayProperties = new List<MemoryArrayProperties>();
        }

    }
}