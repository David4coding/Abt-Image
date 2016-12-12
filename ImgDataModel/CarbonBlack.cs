using System;
using System.ServiceProcess;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace ImgDataModel
{
    public static class CarbonBlack
    {
        public static bool isCbRunning()
        {
            bool result = false;

            try
            {
                ServiceController[] services = ServiceController.GetServices();
                foreach (ServiceController service in services)
                {
                    if (service.ServiceName.Equals("cb"))
                    {
                        result = true;
                        Console.WriteLine("cb Service " + service.ServiceName + " is " + service.Status);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("cb is not running");
            }
            if (!result)
            {
                Console.WriteLine("Carbon Black Service not found");
            }

            return result;
        }

        public static bool isCarbonBlackProcessRunning()
        {
            Process[] processlist = Process.GetProcesses();
            foreach (Process theprocess in processlist)
            {
                if (theprocess.ProcessName.Equals("cb"))
                {

                    Console.WriteLine("Process: {0} ID: {1}", theprocess.ProcessName, theprocess.Id);
                    return true;
                }
            }
            return false;
        }
    }
}
