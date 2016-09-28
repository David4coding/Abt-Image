using Microsoft.Win32;
using System;
namespace VDIDataModel
{
   public static class SilverLight
    {
       private static bool result = false;
        public static bool isInstalled()
        {
            String[] registryValue ;
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

            //    localKey = localKey.OpenSubKey(@"Software\Microsoft\Silverlight\");
            //    registryValue = (int)localKey.GetValue("AllowElevatedTrustAppsInBrowser");

            //    if (registryValue == 1)
            //    {
            //        Console.WriteLine("SilverLigh its installed and the pluging is activated for IE");
            //        result = true;

            //    }
            //}
            //catch (NullReferenceException nre)
            //{
            //    Console.WriteLine(nre.Message);
            //}

            //try
            //{
            //    localKey = localKey.OpenSubKey(@"\SOFTWARE\Wow6432Node\Microsoft\Silverlight");
            //    registryValue = localKey.GetValue("Default").ToString();
            //    if (registryValue.Equals("value not set"))
            //    {
            //        Console.WriteLine("SilverLight its installed but not Available for I.E.");
            //        result = true;
            //    }
            //}
            //catch (NullReferenceException nre)
            //{
            //    Console.WriteLine(nre.Message);
            //}

            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\Microsoft\\Silverlight"))
                {
                    if (key != null)
                    {
                        String o = key.GetValue("Version").ToString();
                        Console.WriteLine(o);
                        result = true;

                    }
                }
            }
            catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
            {
                //react appropriately
                Console.WriteLine(ex.Message);
            }

   

            return result;
        }
    }
}
