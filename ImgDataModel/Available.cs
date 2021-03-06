﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ImgDataModel
{
    class Available
    {
         public static HashSet<string> extensions = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase)
        { ".doc", ".docx", ".xlsx", ".xls",".pdf",".txt", ".pptx",".ppt", ".jpg", ".accdb",".zip",".accdb" };
        public static void DeleteAllFilesInPath(String path)
        {

            if (Directory.Exists(path))
            {
                var files = new DirectoryInfo
               (path)
               .GetFiles()
               .Where(p => extensions.Contains(p.Extension));
                foreach (var file in files)
                {
                    try
                    {
                        file.Attributes = FileAttributes.Normal;
                        File.Delete(file.FullName);
                    }
                    catch (Exception e)
                    {
                       Console.WriteLine("One or more files couldn't be deleted: " + e.Message);
                    }

                }
                Console.WriteLine("Deleting all test files "); // Succes
            }

        }
    }
}
