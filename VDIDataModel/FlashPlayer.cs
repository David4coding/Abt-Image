using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VDIDataModel
{
    public static class FlashPlayer
    {
        private static bool result = false;
        static string registryValue = "empty";
        public static bool isInstalled()
        {
            
            RegistryKey localKey = null;
            if (Environment.Is64BitOperatingSystem)
            {
                localKey = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry64);
            }
            else
            {
                localKey = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry32);
            }

            //try
            //{
            //    localKey = localKey.OpenSubKey(@"SOFTWARE\Wow6432Node\Macromedia\FlashPlayer\");
            //    registryValue = localKey.GetValue("CurrentVersion").ToString();
            //    if (registryValue.Equals("22,0,0,210"))
            //    {
            //        Console.WriteLine("FlashPlayer version : CurrentVersion 22,0,0,210");
            //        result = true;
            //    }
            //}
            //catch (NullReferenceException nre)
            //{
            //    Console.WriteLine(nre.Message);
            //}

            try
            {
                localKey = localKey.OpenSubKey(@"SOFTWARE\Wow6432Node\Macromedia\FlashPlayer\");
                registryValue = localKey.GetValue("Default").ToString();
                //could be changed to Default
                if (registryValue.Equals("value not set"))
                {
                    Console.WriteLine("FlashPlayer is Available: Default");
                    result = true;
                }

            }
            catch (NullReferenceException nre)
            {
                Console.WriteLine(nre.Message);
            }


        


            return result;
        }
    } }

