using System;
using System.IO;

namespace DCOM.Global
{
    public static class FileOperations
    {
        /// <summary>
        /// Verify if a string file Path is actually a Directory
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool IsDirectory(string path)
        {
            // get the file attributes for file or directory
            try
            {
                FileAttributes attr = File.GetAttributes(path); //if not found, then catch statement will return false, i.e. not a directory, and not even a file
                
                if ((attr & FileAttributes.Directory) == FileAttributes.Directory)
                    return true;
                return false;
            }
            catch
            {
                return false;
            }
        }


        /// <summary>
        /// Perform sequential incrementing of the file name. Output string will be the next available integer in the sequence
        /// </summary>
        /// <param name="fullPath"></param>
        /// <returns></returns>
        public static string IncrementFileName(string fullPath)
        {
            int count = 1;
            if (IsDirectory(fullPath))
                return IncrementFolderName(fullPath);

            if (string.IsNullOrEmpty(fullPath))
                throw new ArgumentNullException(nameof(fullPath));

            string fileNameOnly = Path.GetFileNameWithoutExtension(fullPath);
            string extension = Path.GetExtension(fullPath);
            string path = Path.GetDirectoryName(fullPath);

            if (string.IsNullOrEmpty(path))
                throw new ArgumentNullException(nameof(fullPath));

            string tempFileName = string.Format("{0}({1})", fileNameOnly, count);
            string newFullPath = Path.Combine(path, tempFileName + extension);

            while (File.Exists(newFullPath))
            {
                tempFileName = string.Format("{0}({1})", fileNameOnly, count++);
                newFullPath = Path.Combine(path, tempFileName + extension);
            }

            return newFullPath;
        }
        /// <summary>
        /// Resolve a new folder name, using the given <paramref name="fullPath"/> value.
        /// If a pre-existing directory exists, a new name is generated using next incremental index number, starting at 1
        /// </summary>
        /// <param name="fullPath"></param>
        /// <returns></returns>
        public static string IncrementFolderName(string fullPath)
        {
            if (!IsDirectory(fullPath))
                return null;
            int count = 1;
            string newFullPath = fullPath;
            while (Directory.Exists(newFullPath))
            {
                var tempFileName = $"{fullPath}({count++})";
                newFullPath = tempFileName;
            }

            return newFullPath;
        }
        /// <summary>
        /// Update the name of an existing file, as it is version-controlled and/or replaced by a newer file version
        /// </summary>
        /// <param name="fullPath"></param>
        /// <returns></returns>
        public static bool ApplyVersionToFileName(string fullPath)
        {
            string fileNameOnly = Path.GetFileNameWithoutExtension(fullPath);
            string extension = Path.GetExtension(fullPath);
            string path = Path.GetDirectoryName(fullPath);
            string newFullPath = fullPath;

            while (File.Exists(newFullPath))
            {
                string tempFileName = $"{fileNameOnly}_{DateTime.Now.ToString("ddMMMyy_HH_mm_ss")}";
                newFullPath = Path.Combine(path, tempFileName + extension);
            }
            File.Move(fullPath, newFullPath);
            return true;
        }

        /// <summary>
        /// Allow user to call and set location to save their file to
        /// </summary>
        /// <param name="content"></param>
        /// <param name="filePath"></param>
        /// <param name="incrementFileName">Suffix document with next sequential number, if a document of the same name exists. Default will be to overwrite the document.</param>
        /// <returns></returns>
        /// <remarks>If given a directory name, file will be saved using the current DateTime as the filename</remarks>
        public static bool SaveFile(string content, string filePath, bool incrementFileName = false)
        {
            if (String.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException(filePath);
            }

            //see if string is a Full File Path, or just the Directory
            if (IsDirectory(filePath))
            {
                filePath = Path.Combine(filePath, DateTime.Now.ToString("ddMMMyy_HH_mm_ss"));
            }

            string folderPath = Path.GetDirectoryName(filePath);

            //unable to find directory ; throw error
            if (!Directory.Exists(folderPath))
                throw new DirectoryNotFoundException(filePath);

            try
            {
                filePath = incrementFileName ? IncrementFileName(filePath) : filePath;
                File.WriteAllText(filePath, content);
                return true;
            }
            catch (IOException ex)
            {
                throw new IOException("Could not save File to filepath : " + filePath, ex);
            }

        }


        /// <summary>
        /// Allow user to call and set location to save their file to
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="filePath"></param>
        /// <param name="incrementFileName">Suffix document with next sequential number, if a document of the same name exists. Default will be to overwrite the document.</param>
        /// <returns></returns>
        /// <remarks>If given a directory name, file will be saved using the current DateTime as the filename</remarks>
        public static bool SaveFile(Stream stream, string filePath, bool incrementFileName = false)
        {
            if (String.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException(filePath);
            }

            //see if string is a Full File Path, or just the Directory
            if (IsDirectory(filePath))
            {
                filePath = Path.Combine(filePath, DateTime.Now.ToString("ddMMMyy_HH_mm_ss"));
            }

            string folderPath = Path.GetDirectoryName(filePath);

            //unable to find directory ; throw error
            if (!Directory.Exists(folderPath))
                throw new DirectoryNotFoundException(filePath);

            try
            {
                filePath = incrementFileName ? IncrementFileName(filePath) : filePath;
                using (var streamWriter = File.Create(filePath, Int32.MaxValue, FileOptions.None))
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    CopyStream(stream, streamWriter);
                    //stream.CopyTo(streamWriter);
                }
                return true;
            }
            catch (IOException ex)
            {
                throw new IOException("Could not save File to filepath : " + filePath, ex);
            }

        }

        /// <summary>
        /// Copies the contents of input to output. Doesn't close either stream.
        /// </summary>
        private static void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[8 * 1024];
            int len;
            while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, len);
            }
        }
        /// <summary>
        /// Given a current folder location, resolve the complete FilePath for the given <paramref name="relativePath"/>
        /// </summary>
        /// <param name="currentFolder"></param>
        /// <param name="relativePath"></param>
        /// <returns></returns>
        public static string ResolveRelativeFilePath(string currentFolder, string relativePath)
        {
            if (string.IsNullOrEmpty(currentFolder))
                throw new ArgumentNullException(nameof(currentFolder));
            if (string.IsNullOrEmpty(relativePath))
                throw new ArgumentNullException(nameof(relativePath));
            if (!IsDirectory(currentFolder))
                throw new ArgumentException("currentFolder could not be verified as a directory");
            if (Path.IsPathRooted(relativePath))
            {
                return relativePath;    //path is not a relativePath, so return it straight back
            }
            string combined = Path.Combine(currentFolder, relativePath);
            return Path.GetFullPath(combined);
        }

    }
}
