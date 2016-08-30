using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VDISolution
{
    class FlashPlayer
    {
        private bool result = false;
        string registryValue = "empty";
        public bool isInstalled()
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

            try
            {
                localKey = localKey.OpenSubKey(@"SOFTWARE\Wow6432Node\Macromedia\FlashPlayer\");
                registryValue = localKey.GetValue("CurrentVersion").ToString();

            }
            catch (NullReferenceException nre)
            {
                Console.WriteLine(nre.Message);
            }

            if (registryValue.Equals("22,0,0,210") )
            {
                result = true;
            }


            return result;
        }
    } }

