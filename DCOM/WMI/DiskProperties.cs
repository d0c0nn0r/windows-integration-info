using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Net;
using System.Text.RegularExpressions;

namespace WinDevOps.WMI
{
    /// <summary>
    /// Hard Drive properties
    /// </summary>
    /// <remarks>
    /// <see href="https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-logicaldisk"/>
    /// <see href="https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-diskdrivetodiskpartition"/>
    /// <see href="https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-logicaldisktopartition"/>
    /// </remarks>
    public class DiskProperties
    {
        public string Disk { get; internal set; }
        public string Model { get; internal set; }
        public string Partition { get; internal set; }
        public string Description { get; internal set; }
        public string PrimaryPartition { get; internal set; }
        public string VolumeName { get; internal set; }
        public string VolumeSerialNumber { get; internal set; }
        public string FileSystem { get; internal set; }
        public string Drive { get; internal set; }
        /// <summary>
        /// Total Disk size, in bytes
        /// </summary>
        public UInt64 DiskSize { get; internal set; }

        public string DiskSizeString => WMIUtils.FormatBytesToSize((long)DiskSize, 1);

        /// <summary>
        /// Space, in Bytes
        /// </summary>
        public UInt64 FreeSpace { get; internal set; }
        public string FreeSpaceString => WMIUtils.FormatBytesToSize((long)FreeSpace, 1);
        

        public string DiskType { get; internal set; }
        public string SerialNumber { get; internal set; }
        public string BootVolume { get; internal set; }
        public string DriveType { get; internal set; }
        public string DriveLetter { get; internal set; }
        
        /// <summary>
        /// Default internal constructor
        /// </summary>
        internal DiskProperties()
        {
            //default
        }

        private readonly ManagementObject _wmiDiskDrive;
        private readonly ManagementObject _logicalDisk;
        private readonly ManagementObject _partition;
        private readonly ManagementObject _mount;
        internal DiskProperties(ManagementObject diskDrive, ManagementObject logicalDisk, ManagementObject partition) : this()
        {
            _wmiDiskDrive = diskDrive;
            _logicalDisk = logicalDisk;
            _partition = partition;
            this.Disk = diskDrive.GetPropertyValue("Name")?.ToString();
            this.Model = diskDrive.GetPropertyValue("Model")?.ToString();
            this.Partition = partition.GetPropertyValue("Name")?.ToString();
            this.Description = partition.GetPropertyValue("Description")?.ToString();
            this.PrimaryPartition = partition.GetPropertyValue("PrimaryPartition")?.ToString();
            this.VolumeName = logicalDisk.GetPropertyValue("VolumeName")?.ToString();
            this.Drive = logicalDisk.GetPropertyValue("Name")?.ToString();
            this.DiskSize = logicalDisk.GetPropertyValue("Size") != null ? (ulong)logicalDisk.GetPropertyValue("Size") : 0;
            this.FreeSpace = logicalDisk.GetPropertyValue("FreeSpace") != null ? (ulong)logicalDisk.GetPropertyValue("FreeSpace") : 0;
            this.DiskType = "Partition";
            this.SerialNumber = diskDrive.GetPropertyValue("SerialNumber")?.ToString();
            this.FileSystem = "";
            this.BootVolume = "";
        }

        internal DiskProperties(ManagementObject mountPoint) : this()
        {
            _mount = mountPoint;
            this.Disk = mountPoint.GetPropertyValue("Name")?.ToString();
            this.Model = "";
            this.Partition = "";
            this.Description = mountPoint.GetPropertyValue("Caption")?.ToString();
            this.PrimaryPartition = "";
            this.VolumeName = "";
            this.VolumeSerialNumber = "";
            this.FileSystem = mountPoint.GetPropertyValue("FileSystem")?.ToString();
            this.Drive = Regex.Match(mountPoint.GetPropertyValue("Caption").ToString(), "^.:\\\\").Value;
            this.DiskSize = mountPoint.GetPropertyValue("Capacity") !=null ? (ulong)mountPoint.GetPropertyValue("Capacity") : 0;
            this.FreeSpace = mountPoint.GetPropertyValue("FreeSpace") != null ? (ulong)mountPoint.GetPropertyValue("FreeSpace") : 0;
            this.DiskType = "MountPoint";
            this.SerialNumber = mountPoint.GetPropertyValue("SerialNumber")?.ToString();
            this.BootVolume = mountPoint.GetPropertyValue("BootVolume")?.ToString();
            this.DriveType = mountPoint.GetPropertyValue("DriveType")?.ToString();
            this.DriveLetter = mountPoint.GetPropertyValue("DriveLetter")?.ToString();
        }

        internal static List<DiskProperties> Load(ManagementScope scope)
        {
            var returnList = new List<DiskProperties>();
            var diskCollection = WMIUtils.PerformQuery(scope, "SELECT * FROM Win32_DiskDrive")?.ToList();
            if (diskCollection != null)
            {
                //nothing found

                foreach (var diskDrive in diskCollection)
                {
                    var diskDriveClause = diskDrive.GetPropertyValue("DeviceID").ToString().Replace("\\", "\\\\");
                    var paritionQuery =
                        $"ASSOCIATORS OF {{Win32_DiskDrive.DeviceID=\"{diskDriveClause}\"}} WHERE AssocClass = Win32_DiskDriveToDiskPartition";
                    var partitions = WMIUtils.PerformQuery(scope, paritionQuery)?.ToList();
                    if (partitions == null) break;
                    foreach (var partition in partitions)
                    {
                        var logdiskDriveClause = partition.GetPropertyValue("DeviceID").ToString();
                        var logicalDiskQuery =
                            $"ASSOCIATORS OF {{Win32_DiskPartition.DeviceID=\"{logdiskDriveClause}\"}} WHERE AssocClass = Win32_LogicalDiskToPartition";
                        var logicalDisks = WMIUtils.PerformQuery(scope, logicalDiskQuery)?.ToList();
                        if (logicalDisks == null) break;
                        foreach (var logicalDisk in logicalDisks)
                        {
                            returnList.Add(new DiskProperties(diskDrive, logicalDisk, partition));
                        }
                    }
                }
            }

            var mountsCollection =
                WMIUtils.PerformQuery(scope, "SELECT * FROM Win32_Volume WHERE DriveType=3 AND DriveLetter IS NULL")
                    ?.ToList();

            if (mountsCollection == null) return returnList;

            foreach (var mountPoint in mountsCollection)
            {
                returnList.Add(new DiskProperties(mountPoint));
            }

            return returnList;
        }

        public static List<DiskProperties> Load(string computerName = null, NetworkCredential credential = null)
        {
            ManagementScope scope = WMIUtils.GetScope(computerName, credential);
            return Load(scope);
        }
    }
}