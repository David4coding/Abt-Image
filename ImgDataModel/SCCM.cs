﻿using System;
using System.Diagnostics;
using System.ServiceProcess;

namespace ImgDataModel
{
    public static class  SCCM
    {

       public static bool isAgentAvailable()
        {
            bool result = false;

            try
            {
                ServiceController[] services = ServiceController.GetServices();
                foreach (ServiceController service in services)
                {
                    if (service.ServiceName.Equals("CcmExec"))
                    {
                        result = true;
                        Console.WriteLine("SCCM Service "+service.ServiceName + " is " + service.Status);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("CcmExec is not running");
            }

            return result;
        }

    }
}
