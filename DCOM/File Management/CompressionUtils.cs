using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
//using Shell32;
using Ionic.Zip;


namespace WinDevOps.File_Management
{
    public static class CompressionUtils
    {
        public static void CompressFolder(string zipName, string sourceFolder, bool overwrite = false)
        {
            if (zipName == null) throw new ArgumentNullException(nameof(zipName));
            if (sourceFolder == null) throw new ArgumentNullException(nameof(sourceFolder));
            string ifp = Path.GetFullPath(sourceFolder);
            string ofp = Path.GetFullPath(zipName);
            if (Directory.GetParent(ofp) == null)
            {
                throw new ArgumentException($"{nameof(sourceFolder)} value requires a pre-existing parent folder.",
                    sourceFolder);
            }

            if (!Directory.Exists(ifp) || Directory.GetFiles(ifp).Length <= 0)
            {
                throw new ArgumentException($"Folder '{nameof(sourceFolder)}' contains no files, or does not exist.");
            }
            
            if (string.Equals(Directory.GetParent(ofp).FullName, ifp, StringComparison.InvariantCultureIgnoreCase))
            {
                throw new ArgumentException(
                    $"Zip File cannot be located in the {nameof(sourceFolder)} : {nameof(zipName)}={ofp}, {nameof(sourceFolder)}= {ifp}");
            }
            
            if (File.Exists(ofp))
            {
                if (!overwrite) throw new ArgumentException($"Cannot overwrite existing file: {ofp}");
                File.Delete(ofp);
            }

            ////Create an empty zip file
            //byte[] emptyzip = new byte[]{80,75,5,6,0,0,0,0,0,
            //    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0};

            //FileStream fs = File.Create(ofp);
            //fs.Write(emptyzip, 0, emptyzip.Length);
            //fs.Flush();
            //fs.Close();
            //fs = null;
            using (var zip = new ZipFile())
            {
                zip.AddDirectory(ifp);
                zip.Save(ofp);
            }
            //Copy a folder and its contents into the newly created zip file
            //dynamic shell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
            //dynamic source = shell.NameSpace(ifp);
            //dynamic destination = shell.NameSpace(ofp);
            //destination.CopyHere(source.Items(), 20);  //vOptions is binary (20= 4 {No Dialog box} & 16 {Yes to ALL}), https://learn.microsoft.com/en-us/windows/win32/shell/folder-copyhere
            //Shell32.ShellClass sc = new Shell32.ShellClass();
            //Shell32.Folder srcFlder = sc.NameSpace(ifp);
            //Shell32.Folder destFlder = sc.NameSpace(ofp);
            //Shell32.FolderItems items = srcFlder.Items();
            //destFlder.CopyHere(items, 20);

            //Ziping a file using the Windows Shell API 
            //creates another thread where the zipping is executed.
            //This means that it is possible that this console app 
            //would end before the zipping thread 
            //starts to execute which would cause the zip to never 
            //occur and you will end up with just
            //an empty zip file. So wait a second and give 
            //the zipping thread time to get started
            //System.Threading.Thread.Sleep(1000);
        }
    }
}
