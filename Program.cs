using System;
using System.Management;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading;

namespace ContinousRip
{
    class Program
    {
        [DllImport("winmm.dll", CharSet = CharSet.Auto)]
        public static extern int mciSendString(string command, StringBuilder buffer, int bufferSize, IntPtr hwndCallback);
        
        static void Main(string[] args)
        {
            bool runProgram = true;
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            bool isElevated = principal.IsInRole(WindowsBuiltInRole.Administrator);

            if (!isElevated)
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    UseShellExecute = true,
                    WorkingDirectory = Environment.CurrentDirectory,
                    FileName = System.Reflection.Assembly.GetEntryAssembly().Location,
                    Verb = "runas"
                };

                try
                {
                    var process = Process.Start(startInfo);
                    process.WaitForExit();
                    return;
                }
                catch
                {
                    // The user refused to allow privileges elevation.
                    Console.WriteLine("An error occurred while trying to open the drive tray.");
                    return;
                }
            }

            Console.WriteLine("Program starting...");
            
            while (runProgram)
            {
                ManagementObject drive = WaitForDriveTrayClose();
                string driveLetter = drive["Drive"].ToString();
                
                if (IsDiscPresent(drive))
                {
                    string output = RunMakeMKVCON("-r --cache=1 info disc:0");

                    if (drive.Properties["VolumeName"] == null || drive["VolumeName"] == null)
                    {
                        Console.WriteLine("Cannot access VolumeName of the drive. Skipping this run.");
                        continue;
                    }

                    string destinationFolder = Path.Combine(@"C:\Video", drive["VolumeName"].ToString());
                    if (!Directory.Exists(destinationFolder))
                    {
                        Directory.CreateDirectory(destinationFolder);
                    }

                    Console.WriteLine("Ripping disc...");
                    RunMakeMKVCON($"mkv disc:0 all \"{destinationFolder}\"");
                    Console.WriteLine("Done.");

                    OpenDriveTray(driveLetter, ref runProgram);
                }
                else
                {
                    Console.WriteLine("No disc detected. Ejecting drive...");
                    OpenDriveTray(driveLetter, ref runProgram);
                    WaitForDriveTrayClose();
                }

                Thread.Sleep(5000); // Optional delay to give the system some breath
            }
        }

        static ManagementObject WaitForDriveTrayClose()
        {
            Console.WriteLine("Waiting for drive tray to close...");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_CDROMDrive");
            ManagementObject drive = null;

            while (drive == null)
            {
                foreach (ManagementObject obj in searcher.Get())
                {
                    if ((bool)obj["MediaLoaded"])
                    {
                        drive = obj;
                        break;
                    }
                }

                if (drive == null)
                {
                    Thread.Sleep(500);
                }
            }

            Console.WriteLine("Drive tray closed.");
            return drive;
        }

        static bool IsDiscPresent(ManagementObject drive)
        {
            bool discPresent = (bool)drive["MediaLoaded"];
            Console.WriteLine($"Disc present: {discPresent}");
            return discPresent;
        }

        static string RunMakeMKVCON(string parameters)
        {
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = "C:\\Program Files (x86)\\MakeMKV\\makemkvcon64.exe",
                    Arguments = parameters,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };
                Process process = new Process { StartInfo = startInfo };
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();

                return output;
            }
            catch (Exception e)
            {
                Console.WriteLine($"An error occurred while running makemkvcon: {e.Message}");
                throw;
            }
        }

        static void OpenDriveTray(string drive, ref bool runProgram)
        {
            if (drive == null)
            {
                Console.WriteLine("Drive is null, cannot open drive tray.");
                runProgram = false;
                return;
            }

            try
            {
                // Use the ManagementObject to eject the drive
                EjectCD(drive);
                Thread.Sleep(3000); // Give the drive time to open
            }
            catch (ManagementException ex)
            {
                runProgram = false;
                Console.WriteLine($"An error occurred while trying to open the drive tray: {ex.Message}");
            }
        }

        public static void EjectCD(string drive)
        {
            mciSendString($"open {drive} type cdaudio alias drive", null, 0, IntPtr.Zero);
            mciSendString("set drive door open", null, 0, IntPtr.Zero);
        }
    }
}