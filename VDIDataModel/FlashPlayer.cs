using Microsoft.Win32;
using System;

namespace ImgDataModel
{
    public static class FlashPlayer
    {
        public static bool result = false;
        public  static string[] registryValue;
        public static RegistryKey localKey;
        public static string key = "key";
        public static void findRegistryKey()
        {
            try
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
            }
            catch (Exception e)
            {

                Console.WriteLine("Could not find FlashPlayer Registry key "+e.Message);
            }
        }
        public static bool assertFlashPlayer()
        {
            try
            {
                if (localKey != null)
                {
                    localKey = localKey.OpenSubKey(@"SOFTWARE\Wow6432Node\Macromedia\FlashPlayer\");
                    registryValue = localKey.GetValueNames();
                    //could be changed to Default
                    if (registryValue != null)
                    {
                        foreach (var value in registryValue)
                        {
                            key = value.ToString();
                            Console.WriteLine("Registry Key: " + value.ToString());
                            string value1 = localKey.GetValue(key).ToString();
                            Console.WriteLine("Registry Value: " + value1);
                        }
                        result = true;
                    }
                }
                else
                {
                    Console.WriteLine("Registry Key not found : " + key.ToString());
                }
            }
            catch (NullReferenceException nre)
            {
                Console.WriteLine(nre.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            return result;
        }
    }
}

