using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DCOM.Software_Management
{
    public class SoftwareProgram
    {
        public string RegistryKey { get; internal set; }
        public string ComputerName { get; internal set; }

        public string ProgramName { get; internal set; }

        public string Version { get; internal set; }

        public string Publisher { get; internal set; }

        public Dictionary<string, object> Miscellaneous { get; internal set; }

        public SoftwareProgram(string computerName, string programName, string version) : this()
        {
            this.ComputerName = computerName;
            this.ProgramName = programName;
            this.Version = version;
        }

        public SoftwareProgram()
        {
            this.Miscellaneous = new Dictionary<string, object>();
            //default
        }
    }
}
