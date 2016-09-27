using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VDIDataModel
{
   public static class SilverLight
    {
       private static bool result = false;
        public static bool isInstalled()
        {
            int registryValue = 0;
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

                localKey = localKey.OpenSubKey(@"Software\Microsoft\Silverlight\");
                registryValue = (int)localKey.GetValue("AllowElevatedTrustAppsInBrowser");
            }
            catch (NullReferenceException nre)
            {
                Console.WriteLine(nre.Message);
            }

            if(registryValue == 1)
            {
                result = true;
                Console.WriteLine("SilverLigh its installed and the pluging is activated for IE");
            }
            if (registryValue != 1)
            {
                Console.WriteLine("SilverLigh NOT activated for ie");
            }


            return result;
        }
    }
}
