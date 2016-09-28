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
            string Currentdir = Environment.CurrentDirectory;
            string MsbuildDir = Environment.GetEnvironmentVariable("windir")+@"\"+@"Microsoft.NET\Framework64\v4.0.30319";
                 
            List<String> commands = new List<string>();
              Console.WriteLine("Current Directory: " + Currentdir);
            Console.WriteLine("MSBuild Directory:" + MsbuildDir);
            commands.Add("cd "+MsbuildDir);
            commands.Add(@"call msbuild.exe ""C:\Users\"+Environment.UserName + @"\Documents\Visual Studio 2015\Projects\VDISolution\TestRunner\MSBuild Scripts\xunitTask.target");
            commands.Add("cd " + @"C:\Users\"+ Environment.UserName + @"\Documents\Visual Studio 2015\Projects\VDISolution\XmlTransformer\bin\Debug\");
            commands.Add("call XmlTransformer.exe "+
                @"""D:\Jenkins\Workspace\VDI_TEST\VDISolution\TestResults\results.xml"" " +
                @"""C:\Users\frometaguerraj\Documents\Visual Studio 2015\Projects\VDISolution\XmlTransformer\htmlTransform.xsl"" " +
                @"""C:\Users\frometaguerraj\Documents\Visual Studio 2015\Projects\VDISolution\TestResults""");
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
