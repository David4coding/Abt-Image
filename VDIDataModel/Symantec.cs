﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace ImgDataModel
{
    public static class Symantec
    {
        static Process proc = new Process
        {   StartInfo = new ProcessStartInfo
        {
            FileName = @"C:\Program Files (x86)\Symantec\Symantec Endpoint Protection\DoScan.exe",
            Arguments = " /L",
            Domain = "corp.abtassoc.com",
            UseShellExecute = false,
            RedirectStandardError = true,
            RedirectStandardOutput = true
        }
}
        ;

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
                Console.WriteLine("\n\n ERROR: " + e.Message);
            }

            return result;
        }
    }
}
