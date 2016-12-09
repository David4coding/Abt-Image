using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Xml.Xsl;

namespace TestRunner
{
    class Program
    {
        static void Main(string[] args)
        {
            string Currentdir = @"""C:\Users\" + Environment.UserName + @"\Documents\Visual Studio 2015\Projects\AbtImg\AbtRegressionTest";
   
            List<String> commands = new List<string>();
            commands.Add("cd "+ Currentdir);
            Console.WriteLine("Current Directory: " + Currentdir);
            commands.Add(@"call runtests.cmd ");
            executeCommand(commands);
           Console.Read();             
        }
        private static void executeCommand(List<String> commands)
        {
            Process p = new Process();
            ProcessStartInfo info = new ProcessStartInfo("cmd.exe","/k ");
            //info.FileName = "cmd.exe";
            info.RedirectStandardInput = true;
            info.UseShellExecute = false;

            p.StartInfo = info;
            p.Start();

            using (StreamWriter sw = p.StandardInput)
            {
                if (sw.BaseStream.CanWrite)
                {
                    foreach(var command in commands)
                    {
                        sw.WriteLine(command);
                    
                    }
                }
            } 
        }
    }
}
