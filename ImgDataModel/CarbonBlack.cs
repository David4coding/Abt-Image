using System;
using System.ServiceProcess;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
    }
}
