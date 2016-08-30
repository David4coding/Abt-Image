using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace VDISolution
{
    class JAVA
    {
        Process proc = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "java.exe",
                Arguments = " -version",
                UseShellExecute = false,
                RedirectStandardError = true

            }
        };

        public bool CheckVersion()
        {
            bool result = false;
            try
            {
                proc.Start();
                string line = proc.StandardError.ReadLine().Split(' ')[2].Replace("\"", "");
                if (line.Equals("1.6.0_65"))
                {
                    result = true;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
          
            return result;
        }
        public void readOutPutLines()

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
