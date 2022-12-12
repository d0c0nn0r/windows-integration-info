using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CsvHelper.Configuration;
using DCOM.Software_Management;

namespace win_int_info.Serialization
{
    internal sealed class SoftwareProgramMap : CsvClassMap<SoftwareProgram>
    {
        public SoftwareProgramMap()
        {
            this.AutoMap();
            this.Map(m => m.ComputerName).Name("PSComputerName");
            this.Map(m => m.Miscellaneous).Name("Miscellaneous").TypeConverter<EnumerableConverter>();
        }
    }
}
