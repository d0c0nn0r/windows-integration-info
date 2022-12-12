using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace WinDevOps.File_Management
{
    public static class FileUtils
    {
        public static void CopyFile(string target, string destination, bool overwrite=true)
        {
            if (!File.Exists(target)) throw new FileNotFoundException("File was not found", target);
            if (File.Exists(destination))
            {
                if (!overwrite)
                    throw new ArgumentException(
                        $"File name already exists: {destination}. Provide unique name, or set overwrite=true");
            }
            File.Copy(target, destination, overwrite);
        }
    }
}
