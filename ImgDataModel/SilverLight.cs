using Microsoft.Win32;
using System;
namespace ImgDataModel
{
   public static class SilverLight
    {
       private static bool result = false;
        public static String[] registryValue;
        public static RegistryKey localKey = null;
        public static void findRegistry()
        {
           
            if (Environment.Is64BitOperatingSystem)
            {
                localKey = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry64);
            }
            else
            {
                localKey = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, RegistryView.Registry32);
            }
         
           
        }

        public static bool checkSilverLightVersion()
        {
            try
            {
                using (RegistryKey regkey = Registry.LocalMachine.OpenSubKey("Software\\Wow6432Node\\Microsoft\\Silverlight"))
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
                            string[] value1 = localKey.GetValueNames();
                            //print values
                            foreach (var item in value1)
                            {
                                Console.WriteLine("Registry Value: " + item);
                            }
                           
                        }
                       
                    }
                    else
                    {
                        Console.WriteLine("Registry Value not found, instead "+ registryValue.ToString());
                    }
                }
            }
            catch (Exception ex)  //just for demonstration...it's always best to handle specific exceptions
            {
                //react appropriately
                Console.WriteLine("Couldnt find the Silverlight registry");
            }
            return result;
        }
    }
}
