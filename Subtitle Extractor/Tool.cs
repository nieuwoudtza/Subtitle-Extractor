using IWshRuntimeLibrary;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Threading.Tasks;

namespace Subtitles_Extractor
{
    public static class Tool
    {
        static string applicationPath = Assembly.GetExecutingAssembly().Location;
        static string sendTo = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Microsoft\Windows\SendTo\Subtitles Extractor.lnk";

        public static void CreateShortcut()
        {
            WshShell wsh = new WshShell();
            IWshShortcut shortcut = wsh.CreateShortcut(sendTo);
            shortcut.Arguments = "";
            shortcut.TargetPath = applicationPath;
            shortcut.WindowStyle = 1;
            shortcut.IconLocation = applicationPath;
            shortcut.Save();
        }

        public static bool ShortcutExists()
        {
            return System.IO.File.Exists(sendTo);
        }

        public static void RemoveShortcut()
        {
            System.IO.File.Delete(sendTo);
        }

        public static List<string> SearchDirectory(string dir, string[] searchPattern, bool subdirectories, int levels = 0, int currentLevel = 0)
        {
            List<string> files = new List<string>();
            try
            {
                if (searchPattern.Length > 0)
                {
                    foreach (string pattern in searchPattern)
                        foreach (string file in Directory.GetFiles(dir, pattern))
                            files.Add(file);
                }
                else
                {
                    foreach (string file in Directory.GetFiles(dir))
                        files.Add(file);
                }

                if (subdirectories && ++currentLevel != levels)
                    foreach (string directory in Directory.GetDirectories(dir))
                        files.AddRange(SearchDirectory(directory, searchPattern, subdirectories, levels, currentLevel));
            }
            catch { }

            return files;
        }

        public static void ExtractToDirectory(byte[] source, string destinationDirectory, bool overwrite)
        {
            string Temp = Path.GetTempPath() + Path.GetRandomFileName() + ".zip";
            System.IO.File.WriteAllBytes(Temp, source);

            ZipArchive archive = new ZipArchive(new FileStream(Temp, FileMode.Open));


            foreach (ZipArchiveEntry file in archive.Entries)
            {
                string completeFileName = Path.Combine(destinationDirectory, file.FullName);
                if (overwrite || !System.IO.File.Exists(completeFileName))
                {
                    string directory = Path.GetDirectoryName(completeFileName);

                    if (!Directory.Exists(directory))
                        Directory.CreateDirectory(directory);

                    if (file.Name != "")
                        file.ExtractToFile(completeFileName, true);
                }
            }

            TryDelete(Temp, 10);
        }

        public static async void TryDelete(string path, int attempts)
        {
            bool success = false;

            while (!success)
            {
                try
                {
                    System.IO.File.Delete(path);
                    success = true;
                }
                catch
                {
                    await Task.Delay(1000);
                }
            }
            
        }
    }
}
