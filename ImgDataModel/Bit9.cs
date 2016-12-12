using System;
using System.ServiceProcess;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace ImgDataModel
{
    public static class Bit9
    {
        public static bool isParityRunning()
        {
            bool result = false;

            try
            {
                ServiceController[] services = ServiceController.GetServices();
                foreach (ServiceController service in services)
                {
                    if (service.ServiceName.Equals("parity"))
                    {
                        result = true;
                        Console.WriteLine("parity Service " + service.ServiceName + " is " + service.Status);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("parity is not running");
            }
            if (!result)
            {
                Console.WriteLine("Bit9 Service not found");
            }

            return result;
        }

        public static bool isBit9ProcessRunning()
        {
            Process[] processlist = Process.GetProcesses();
            foreach (Process theprocess in processlist)
            {
                if (theprocess.ProcessName.Equals("parity"))
                {

                    Console.WriteLine("Process: {0} ID: {1}", theprocess.ProcessName, theprocess.Id);
                    return true;
                }
            }
            return false;
        }

    }
}
