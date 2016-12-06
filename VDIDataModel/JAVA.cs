using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace ImgDataModel
{
    public static class JAVA
    {
        static Process proc = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "java.exe",
                Arguments = " -version",
                UseShellExecute = false,
                RedirectStandardError = true

            }
        };

        public static bool CheckVersion()
        {
            bool result = false;
            try
            {
                proc.Start();
                string line = proc.StandardError.ReadLine().Split(' ')[2].Replace("\"", "");
                //Console.WriteLine(line);
                if (line.Equals("1.8.0_101"))  //1.6.0_65 for VDI, for server 1.8.0_102 visual studio : 1.7.0_71
                {
                    Console.WriteLine("JAVA Version : 1.6.0_65");
                    result = true;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
          
            return result;
        }
        public static void readOutPutLines()

        {

            try
            {
                proc.Start();
                while (!proc.StandardError.EndOfStream)
                {
                    var line = proc.StandardError.ReadLine().Split(' ')[2].Replace("\"", "");
                    Console.WriteLine("line : " + line);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("\n\n ERROR: " + e.Message);

            }


            Console.WriteLine("\n\n Press any key to exit.");
            Console.ReadLine();
            //  proc.Close();
            //or any other statements for that matter
        }

    }
}
