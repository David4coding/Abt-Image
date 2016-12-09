using System;
using Microsoft.Win32;
using System.Diagnostics;
using System.ServiceProcess;

namespace ImgDataModel
{

    public static class Symantec
    {
        private static bool result = false;
        public static String[] registryValue;
        public static RegistryKey localKey = null;
        static Process proc = new Process

        {
            StartInfo = new ProcessStartInfo
            {
                FileName = @"C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\DoScan.exe",
                Arguments = " /L",
                Domain = "corp.abtassoc.com",
                UseShellExecute = false,
                RedirectStandardError = true,
                RedirectStandardOutput = true
            }
        };
        

        public static bool isSymantecActive()
        {
            bool result= false;
            try
            {
                proc.Start();
                while (!proc.StandardOutput.EndOfStream)
                {
                    string output = proc.StandardOutput.ReadLine();
                    string strOutput = proc.StandardError.ReadLine();
                    if (output != null && output.Length > 2)
                    {
                        //  Console.WriteLine("complete output : " + output);
                        result = output.Contains("Active Scan Upon Startup");
                        Console.WriteLine("Symantec Active Scan Upon Startup: " + result);
                       
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error while checking Symantec Active Scan");
            }
            return result;
        }

        public static bool isEncriptionServiceRunning()
        {
            bool result = false;

            try
            {
                ServiceController[] services = ServiceController.GetServices();
                foreach (ServiceController service in services)
                {
                    if (service.ServiceName.Equals("PGPtray"))
                    {
                        result = true;
                        Console.WriteLine("PGPtray Service " + service.ServiceName + " is " + service.Status);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("pgptray is not running");
            }

            return result;
        }

        public static bool isEncriptionProcessRunning()
        {
            Process[] processlist = Process.GetProcesses();
            foreach (Process theprocess in processlist)
            {
                if (theprocess.ProcessName.Equals("PGPtray"))
                {

                    Console.WriteLine("Process: {0} ID: {1}", theprocess.ProcessName, theprocess.Id);
                    return true;
                }
            }
            return false;
        }

        public static bool findEncriptionRegistry()
        {
            try
            {


                if (Environment.Is64BitOperatingSystem)
                {
                    localKey = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry64);
                }
                else
                {
                    localKey = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry32);
                }


                using (RegistryKey regkey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Wow6432Node\\PGP Corporation\\PGP"))
                {
                    registryValue = regkey.GetValueNames();
                    //could be changed to Default
                    if (registryValue != null)
                    {
                        foreach (var value in registryValue)
                        {
                            string key = value.ToString();
                           Console.WriteLine("Registry Key: " + value.ToString());
                            result = true;
                            //string[] value1 = localKey.GetValueNames();
                            ////print values
                            //foreach (var item in value1)
                            //{
                            //    if (item.Equals("PGPSTAMP"))
                            //    {
                            //        Console.WriteLine("Registry Value: " + item);
                            //    }
                            //}
                        }
                    }
                    else
                    {
                        Console.WriteLine("Registry Value not found, instead " + registryValue.ToString());
                    }
                }
            }
            catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
            {
                //react appropriately
                Console.WriteLine("Couldnt find the Symantec Desktop Encryption registry");
            }
            return result;
        }

    }



}
