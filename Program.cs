using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace TestApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string[] serviceNames = { "WyseDeviceAgent", "WysePlatformService" }; 
            try
            {
                //reg creation 
                string regpath = "SOFTWARE\\WNT";
                RegistryKey key;
                key=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(regpath, true);
                if (key != null)
                {
                    object value = key.GetValue("MerlinUpdate");
                    if (value!=null)
                    {
                        ExecutePowerShellScript("C:\\Windows\\Setup\\Tools\\Merlin\\LastEntry.ps1");

                        Thread.Sleep(20000);
                        foreach (string serviceName in serviceNames)
                        {
                            DisableService(serviceName, "auto");
                        }

                        //Restart Device
                        Process.Start("shutdown", "/r /t 0 /f");
                    }
                    else
                    {
                        foreach (string serviceName in serviceNames)
                        {
                            StartStopService(serviceName, "stop");
                            DisableService(serviceName, "disabled");
                        }
                        try
                        {
                            var Add = AssignLetterUsingDiskpart("4", "F:", "assign");
                            System.Threading.Thread.Sleep(500);
                            if (Add)
                            {
                                CopyFiles("C:\\Windows\\Setup\\Tools\\Merlin\\DELL", "F:\\EFI\\DELL");
                                CopyFiles("C:\\Windows\\Setup\\Tools\\Merlin\\Merlin", "F:\\EFI\\Merlin");
                            }
                            var Remove = AssignLetterUsingDiskpart("4", "F:", "remove");
                        }
                        catch (Exception ex) { }

                        key.SetValue("MerlinUpdate", "1", RegistryValueKind.String);
                        ExecutePowerShellScript("C:\\Windows\\Setup\\Tools\\Merlin\\FirstEntry.ps1");
                    }                    
                }
            }
            catch (Exception ex)
            {

            }           
        }

        static void ExecutePowerShellScript(string filePS)
        {
            try
            {
                // Path to PowerShell executable
                string powerShellExe = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe";

                // Path to the PowerShell script
                string scriptPath = filePS;

                // Create the PowerShell process start info
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = powerShellExe,
                    Arguments = $"-File \"{scriptPath}\"",
                    Verb = "runas", // Run as administrator
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using (Process process = new Process { StartInfo = psi })
                {
                    // Start the process
                    process.Start();

                    // Read the output and error streams
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();

                    // Display the output and error (you can modify this part based on your needs)
                    Console.WriteLine("Output:\n" + output);
                    //Console.WriteLine("Error:\n" + error);

                    // Wait for the process to exit
                    process.WaitForExit();
                }

                System.Threading.Thread.Sleep(5000);
            }
            catch (Exception ex)
            {
                //Console.WriteLine("Error: " + ex.Message);
            }
        }

        /// <summary>
        /// This function will be used to handling the Files and Folder Copy 
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="destinationPath"></param>
        static void CopyFiles(string sourcePath, string destinationPath)
        {
            try
            {
                if (!Directory.Exists(destinationPath))
                    Directory.CreateDirectory(destinationPath);

                DirectoryInfo sourceDir = new DirectoryInfo(sourcePath);
                DirectoryInfo destinationDir = new DirectoryInfo(destinationPath);

                //Files Copy
                foreach (FileInfo file in sourceDir.GetFiles())
                {
                    string tempPath = Path.Combine(destinationDir.FullName, file.Name);
                    file.CopyTo(tempPath, true);
                }
                //Directory Copy
                foreach (DirectoryInfo dir in sourceDir.GetDirectories())
                {
                    string tempPath = Path.Combine(destinationDir.FullName, dir.Name);
                    CopyFiles(dir.FullName, tempPath);
                }

                //Folder Deletion
                Directory.Delete(sourcePath,true);
            }
            catch (Exception ex)
            { }
        }

        /// <summary>
        /// This function is created for handling the Assigning/Removing letter to Vol
        /// </summary>
        /// <param name="volumeName">1/2/3/4</param>
        /// <param name="driveLetter">F:</param>
        static bool AssignLetterUsingDiskpart(string volumeName, string driveLetter,string addremove)
        {
            bool res = false;
            try
            {
                string diskpartScript = $"select volume \"{volumeName}\"\n {addremove} letter={driveLetter}";

                // Create the PowerShell process start info
                ProcessStartInfo processInfo = new ProcessStartInfo
                {
                    FileName = "diskpart.exe",
                    Verb = "runas", // Run as administrator
                    RedirectStandardInput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(processInfo))
                {
                    if (process != null)
                    {
                        using (StreamWriter sw = process.StandardInput)
                        {
                            if (sw != null)
                            {
                                sw.WriteLine(diskpartScript);
                            }
                        }
                        process.WaitForExit();
                    }
                }
                res = true;
            }
            catch (Exception ex)
            { }
            return res;
        }

       /// <summary>
       /// Starting and stopping service
       /// </summary>
       /// <param name="serviceName"></param>
       /// <param name="startstop"></param>
        static void StartStopService(string serviceName ,string startstop)
        {
            try
            {
                using (ServiceController serviceController = new ServiceController(serviceName))
                {
                    if(startstop=="stop")
                    {
                        if (serviceController.Status != ServiceControllerStatus.Stopped)
                        {
                            serviceController.Stop();
                            serviceController.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(10));
                        }
                    }
                    else
                    {
                        if (serviceController.Status != ServiceControllerStatus.Running)
                        {
                            serviceController.Start();
                            serviceController.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(10));
                        }
                    }
                   
                }
            }
            catch(Exception ex)
            {

            }            
        }

        /// <summary>
        /// Disabling the service
        /// </summary>
        /// <param name="serviceName"></param>
        /// <param name="status">disabled/auto</param>
        static void DisableService(string serviceName , string status)
        {
            try
            {
                using (System.Diagnostics.Process process = new System.Diagnostics.Process())
                {
                    process.StartInfo.FileName = "sc.exe";
                    process.StartInfo.Arguments = $"config {serviceName} start={status}";
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.CreateNoWindow = true;
                    process.Start();
                    process.WaitForExit();
                }
            }
            catch (Exception ex)
            {

            }            
        }
    }
} 
