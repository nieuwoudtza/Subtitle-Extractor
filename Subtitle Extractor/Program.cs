using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

namespace Subtitles_Extractor
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Title = "Subtitle Extractor";

            if (args.Length == 0)
            {
                if (!Tool.ShortcutExists())
                {
                    Tool.CreateShortcut();
                    Console.Write("'Send-To' Shortcut Created");
                }
                else
                {
                    Tool.RemoveShortcut();
                    Console.Write("'Send-To' Shortcut Removed");
                }

                Thread.Sleep(2000);
                Environment.Exit(0);
            }


            List<string> Args = new List<string>();

            for (int i = 0; i < args.Length; i++)
            {
                if (File.Exists(args[i]))
                    Args.Add(args[i]);
                else if (Directory.Exists(args[i]))
                    Args.AddRange(Tool.SearchDirectory(args[i], new string[] { "*.*" }, true, 2));
            }

            if (Args.Count == 0)
                Environment.Exit(0);

            Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory);

            if (!File.Exists("ffmpeg.exe") || !File.Exists("SE361\\SubtitleEdit.exe"))
                Tool.ExtractToDirectory(Properties.Resources.Dependencies, AppDomain.CurrentDomain.BaseDirectory, true);

            Console.Write("Language (eng - default): ");

            string lang = Console.ReadLine().ToLower();

            if (string.IsNullOrEmpty(lang))
                lang = "eng";

            List<Thread> Extracts = new List<Thread>();

            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Selected language: " + lang);
            Console.WriteLine();

            for (int i = 0; i < Args.Count; i++)
            {
                Process proc = new Process();
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.CreateNoWindow = true;
                proc.StartInfo.FileName = "ffmpeg";
                proc.StartInfo.Arguments = "-i \"" + Args[i] + "\"";
                proc = Process.Start(proc.StartInfo);

                string[] output = proc.StandardError.ReadToEnd().ToLower().Split(new string[] { "\n", "\r" }, StringSplitOptions.RemoveEmptyEntries);

                proc.WaitForExit();

                for (int j = 0; j < output.Length; j++)
                {
                    string line = output[j].Trim();

                    if (line.Contains("("+ lang + "): subtitle: subrip") || line.Contains("(" + lang + "): subtitle: mov_text"))
                    {
                        line = line.Substring("stream #".Length);
                        string stream = line.Substring(0, line.IndexOf("("));

                        Extracts.Add(Extract_Text(Args[i], stream));
                        break;
                    }
                    else if (line.Contains("(" + lang + "): subtitle: hdmv_pgs_subtitle"))
                    {
                        line = line.Substring("stream #".Length);
                        string stream = line.Substring(0, line.IndexOf("("));

                        Extracts.Add(Extract_Images(Args[i], stream));
                        break;
                    }
                }

                for (int j = 0; j < output.Length; j++)
                {
                    string line = output[j].Trim();

                    if (line.Contains("(und): subtitle: subrip") || line.Contains("(und): subtitle: mov_text"))
                    {
                        line = line.Substring("stream #".Length);
                        string stream = line.Substring(0, line.IndexOf("("));

                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine("Language 'und' (undetermined) used for file "+ Args[i]);
                        Console.WriteLine();
                        Extracts.Add(Extract_Text(Args[i], stream));
                        break;
                    }
                    else if (line.Contains("(und): subtitle: hdmv_pgs_subtitle"))
                    {
                        line = line.Substring("stream #".Length);
                        string stream = line.Substring(0, line.IndexOf("("));

                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine("Language 'und' (undetermined) used for file " + Args[i]);
                        Console.WriteLine();
                        Extracts.Add(Extract_Images(Args[i], stream));
                        break;
                    }
                }

                if (Extracts.Count == 0)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("No valid subtitles found for file " + Args[i]);
                    Console.WriteLine();
                }
                
            }

            for (int k = 0; k < Extracts.Count; k++)
            {
                if (Extracts[k].IsAlive)
                    Extracts[k].Join();

                Console.Title = "Subtitle Extractor | " + ((double)(k + 1) / Extracts.Count * 100) + "%";
            }

            Console.Title = "Subtitle Extractor | Done";
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write("All tasks completed");
            Console.ReadKey();
        }

        static Thread Extract_Text(string input, string stream)
        {
            Thread extract = new Thread(() =>
            {
                ProcessStartInfo psi = new ProcessStartInfo("ffmpeg.exe", "-i \"" + input + "\"" + " -map " + stream + " \"" + Path.ChangeExtension(input, ".srt") + "\" -y");
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                psi.RedirectStandardError = true;

                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("Extracting .srt from " + input);
                Console.WriteLine();

                Process _extract = new Process();
                _extract.StartInfo = psi;
                _extract.Start();
                string Error = GetError(_extract.StandardError.ReadToEnd());
                _extract.WaitForExit();

                if (Error.StartsWith("video:0kB audio:0kB subtitle:") && Error.EndsWith("%"))
                    Error = string.Empty;

                if (!string.IsNullOrEmpty(Error))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Couldn't extract from '" + input + "': " + Error);
                    Console.WriteLine();
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Finished extraction from " + input);
                    Console.WriteLine();
                }
                
            });
            extract.Start();

            return extract;
        }

        static Thread Extract_Images(string input, string stream)
        {
            Thread extract = new Thread(() =>
            {
                string SupFile = Path.ChangeExtension(input, ".sup");

                ProcessStartInfo psi = new ProcessStartInfo("ffmpeg.exe", "-i \"" + input + "\" -map " + stream + " -c:s copy \"" + SupFile + "\" -y");
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                psi.RedirectStandardError = true;

                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("Extracting .sup from " + input);
                Console.WriteLine();

                Process _extract = new Process();
                _extract.StartInfo = psi;
                _extract.Start();
                string Error = GetError(_extract.StandardError.ReadToEnd());
                _extract.WaitForExit();
                //string Error = GetError(_extract.StandardError.ReadToEnd());

                FileInfo fi = new FileInfo(SupFile);

                if (!fi.Exists || (!string.IsNullOrEmpty(Error) && fi.Length == 0))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Couldn't extract from '" + input + "': " + Error);
                    Console.WriteLine();
                    return;
                }

                string TempFile = Path.GetTempFileName();
                File.Delete(TempFile);
                TempFile = Path.ChangeExtension(TempFile, ".bat");

                string SubEditPath = AppDomain.CurrentDomain.BaseDirectory + "SE361\\SubtitleEdit.exe";
                File.WriteAllText(TempFile, "\"" + SubEditPath + "\" /convert \"" + SupFile + "\"" + " SubRip /fps:25");

                psi = new ProcessStartInfo(TempFile);
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                psi.RedirectStandardError = true;

                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.WriteLine("Performing OCR on " + SupFile);
                Console.WriteLine();

                Process _convert = new Process();
                _convert.StartInfo = psi;
                _convert.Start();
                Error = GetError(_convert.StandardError.ReadToEnd());
                _convert.WaitForExit();

                

                try
                {
                    if (File.Exists(SupFile))
                        File.Delete(SupFile);

                    if (File.Exists(TempFile))
                        File.Delete(TempFile);
                }
                catch { }

                if (!string.IsNullOrEmpty(Error))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Couldn't extract from '" + input + "': " + Error);
                    Console.WriteLine();
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Finished extraction from " + input);
                    Console.WriteLine();
                }
            });
            extract.Start();

            return extract;
        }

        static string GetError(string Output)
        {
            if (Output.Length > 0)
            {
                string[] Error = Output.Split(new string[] { "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                return Error.Last();
            }

            return "";
        }
    }
}
