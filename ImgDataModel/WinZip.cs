using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ImgDataModel
{
   public static class WinZip
    {
        //path for the test to be written
        public static string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                                 + @"\VDI_FRAMEWORK_RUNNER\";
        public static string fileName = "VDI.txt";

        public static void WritteBinaryFile()
        {
           

            try
            {   //bynary blocks set up
                const int blockSize = 1024; //1kb
                const int blocksPerMb = (1024 * 100) / blockSize; // 100kb
                //random data
                byte[] data = new byte[blockSize];
                Random rng = new Random();
              
                // the code that you want to measure comes here
                using (FileStream stream = File.OpenWrite
                    (path+fileName))
                {
                    for (int i = 0; i < 1 * blocksPerMb; i++)
                    {
                        rng.NextBytes(data);
                        stream.Write(data, 0, data.Length);
                    }
                }
                
            }
            catch (IOException e)
            {
                Console.WriteLine(e.Message + "\n Cannot write " + path+fileName + ".");
                return;
            }

        }

        public static bool addToZip()
        {
            bool result = false;
            try
            {
                using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile())
                {
                    zip.AddFile(path+fileName);
                    zip.Save(path+"VDIZip.zip");
                    Console.WriteLine("Zip File Created Sucessfully:");
                    
                }
            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return result;
        }


        /// <summary>
        /// searches for all the permited formats and delete those files.
        /// </summary>
       
        public static bool CheckZip(string zip)
        {
            if (File.Exists(zip))
            {
                Available.DeleteAllFilesInPath(path);
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}
