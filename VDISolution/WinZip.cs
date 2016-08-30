using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace VDISolution
{
    class WinZip
    {

     public bool isAvailable()
        {
            bool result = false;
            try
            {
                using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
                {
                    // add this map file into the "images" directory in the zip archive
                    zip.AddFile(@"C:\QA_Development_Scripts\VDISolution\ZipTestImage.jpg", "images");
                    // add the report into a different directory in the archive
                    zip.AddFile(@"C:\QA_Development_Scripts\VDISolution\ZipTestFile.docx", "files");
                    zip.AddFile(@"C:\QA_Development_Scripts\VDISolution\ZipReadMe.txt");
                    zip.Save(@"C:\QA_Development_Scripts\VDISolution\ZipSuccess.zip");
                    result = CheckZip(@"C:\QA_Development_Scripts\VDISolution\ZipSuccess.zip");
                }
            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return result;
        }
        public bool CheckZip(string zip)
        {
            if (File.Exists(zip))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}
