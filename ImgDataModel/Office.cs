
using Microsoft.Office.Interop.Word;
using System;
using System.DirectoryServices;
using Microsoft.Office.Interop.Access;
using System.IO;
using System.Reflection;


namespace ImgDataModel
{
    public static class Office
    {
        //path for the test to be written
        private static string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                                 + @"\VDI_FRAMEWORK_RUNNER\";
        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception Occured while releasing object ");
            }
            finally
            {
                GC.Collect();
            }
        }
        public static class OutlookWrapper
        {
            private static Microsoft.Office.Interop.Outlook.Application outlook;
            public static string OutlookUser;
            public static string WindowsUser;
            public static DirectoryEntry de;
            public static bool result = false;
            public static void hasActiveDirectoryCredentials()
            {
                outlook = new Microsoft.Office.Interop.Outlook.Application();
                OutlookUser = outlook.Application.Session.CurrentUser.Name;
                de = new DirectoryEntry("WinNT://" + Environment.UserDomainName + "/" + Environment.UserName);
                WindowsUser = de.Properties["fullName"].Value.ToString();

            }
            public static bool assertOutlook()
            {
                if (OutlookUser == WindowsUser)
                {
                    Console.WriteLine("Outlook User: " + WindowsUser);
                    result = true;
                }
                return result;
            }
        }
        public static class WordWrapper
        {
            /// <summary>
            /// creates a word file .docx
            /// </summary>
            /// <param name="fileName">template file for the Copy funct</param>
      
            public static Microsoft.Office.Interop.Word.Application oWord;
            public static _Document oDoc;
            public static Paragraph oPara1;

            private static object  oMissing = System.Reflection.Missing.Value;
            private static object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
            public static bool result = false;
            public static string filename = @"VDI.docx";

            public static void CreateWord()
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                try
                {
                    //Start Word and create a new document.
                    oWord = new Microsoft.Office.Interop.Word.Application();
                    oWord.Visible = false;
                    oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Word Doc could not be created at this time ");

                }
            }
            public static void addParagraph()
            {
                try
                {
                    //Insert a paragraph at the beginning of the document.
                    oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
                    oPara1.Range.Text = "VDI MICROSOFT TEST";
                    oPara1.Range.Font.Bold = 1;
                    oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
                    oPara1.Range.InsertParagraphAfter();
                }catch (Exception e)
                {
                    Console.WriteLine("The paragraph could not be added at this time ");

                }
            }
            public static void saveFile()
            {
                try
                {
                    oDoc.SaveAs(path + filename);
                    oDoc.Close();
                    oWord.Quit();
                    //realease from memory and call the GC
                    releaseObject(oDoc);
                    releaseObject(oWord);
                }catch(Exception e)
                {
                    Console.WriteLine("Error saving the Word Doc: " + filename);

                }
            }
            public static bool assertWordDoc()
            {
                try
                {
                    //asserts the file exists
                    if (File.Exists(path + "VDI.docx"))
                    {
                        result = true;
                        Console.WriteLine("Word file saved as: " + filename);
                        Available.DeleteAllFilesInPath(path);
                    }
                  
                }catch(Exception e)
                {
                    Console.WriteLine("file " + filename+ " not found");
                }
                return result;
            }
        }
        public static class ExcelWrapper
        {
            public static Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            public static Microsoft.Office.Interop.Excel.Workbook excelWorkbook = ExcelApp.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            public static Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Worksheets[1];

            private static object oMissing = System.Reflection.Missing.Value;
            private static object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
            public static bool result = false;
            public static string fileName = @"VDI.xlsx";

            public static void CreateExcel()
            {
               

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                try
                {
                    if (ExcelApp == null)
                    {
                        Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                        
                    }
                    ExcelApp.Visible = false;

                    if (excelWorksheet == null)
                    {
                        Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
                    }
                    
                }
                catch (Exception e)
                {
                    Console.WriteLine(" Opening Excel, \n Add new content to Excel, \n Saving Failed: " + fileName + "\n" + e.Message); // Success

                }
            
            }

            public static void addRows()
            {
                // Select the Excel cells, in the range c1 to c7 in the worksheet.
                Microsoft.Office.Interop.Excel.Range aRange = excelWorksheet.get_Range("C1", "C7");

                if (aRange == null)
                {
                    Console.WriteLine("Could not get a range. Check to be sure you have the correct versions of the office DLLs.");
                }

                // Fill the cells in the C1 to C7 range of the worksheet with the number 6.
                Object[] args = new Object[1];
                args[0] = 6;
                aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);

                // Change the cells in the C1 to C7 range of the worksheet to the number 8.
                aRange.Value2 = 21;
            }

            public static void saveExcel()
            {
                try
                {
                    excelWorkbook.SaveAs(path + fileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing,
                    Type.Missing, true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

                    excelWorkbook.Close();// change missValue to null
                    ExcelApp.Quit();
                    //relese from memory
                    releaseObject(excelWorksheet);
                    releaseObject(excelWorkbook);
                    releaseObject(ExcelApp);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error saving the Excel file: "+fileName);
                }
            }

            public static bool assertExcel()
            {
                try
                {
                    //asserts the file exists
                    if (File.Exists(path + "VDI.xlsx"))
                    {
                        result = true;
                        Console.WriteLine("Excel file saved as: " + fileName);
                    }
                    Available.DeleteAllFilesInPath(path);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Excel file " + fileName + " not found");
                }
                return result;
            }
        }
        public static class PowerPointWrapper
        {
            public static Microsoft.Office.Interop.PowerPoint.Application ppApp = new Microsoft.Office.Interop.PowerPoint.Application();
            public static Microsoft.Office.Interop.PowerPoint.Presentations ppPresens ;
            public static Microsoft.Office.Interop.PowerPoint.Presentation objPres ;
            public static Microsoft.Office.Interop.PowerPoint.Slides objSlides ;
            public static Microsoft.Office.Interop.PowerPoint.Slide objSlide ;
            public static Microsoft.Office.Interop.PowerPoint.TextRange objTextRng;

            public static bool result = false;
            public static string fileName = "VDI.pptx";

            public static bool CreatePowerPoint()
            {
             
                try
                {   //create powerpoint
                    ppPresens = ppApp.Presentations;
                    objPres = ppPresens.Add();
                    objSlides = objPres.Slides;

                }
                catch (Exception e)
                {
                    Console.WriteLine("Creating PowerPoint Failed: " + fileName + "\n" + e.Message);

                }
                return result;
            }

            public static void addText()
            {
                try
                {
                    //addind text to the slides
                    objSlide = objSlides.Add(1, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
                    objTextRng = objSlide.Shapes[1].TextFrame.TextRange;

                    objTextRng.Text = "VDI POWER POINT TEST";
                    objTextRng.Font.Name = "Comic Sans MS";
                    objTextRng.Font.Size = 48;
                }
                catch (Exception e)
                {
                    Console.WriteLine("Adding Text to " + fileName +" failed");

                }
            }

            public static void savePowerPoint()
            {
                try
                {
                    //saving powerpoint
                    objPres.SaveAs(path + fileName, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault);
                    
                    ppApp.Quit();

                    //release memory
                    releaseObject(objTextRng);
                    releaseObject(objSlide);
                    releaseObject(objSlides);
                    releaseObject(objPres);
                    releaseObject(ppPresens);
                    releaseObject(ppApp);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error saving " + fileName );

                }
            }

            public static bool assertPowerPoint()
            {

                try
                {
                    //asserts the file exists
                    if (File.Exists(path + "VDI.pptx"))
                    {
                        result = true;
                        Console.WriteLine("PowerPoint file saved as: " + fileName);
                    }


                    Available.DeleteAllFilesInPath(path);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Assert PowerPoint Failed: " + fileName + "\n" + e.Message);

                }
                return result;
            }

        }
        public static class AccessWrapper
        {
            public static Microsoft.Office.Interop.Access.Application AccessAPP = new Microsoft.Office.Interop.Access.Application();
            public static bool result = false;
            public static string fileName = "VDI.accdb";
            public static void CreateAccess()
            {
               
                try
                {
                    //create access database
                    AccessAPP.NewCurrentDatabase(path+fileName,
                                         Microsoft.Office.Interop.Access.AcNewDatabaseFormat.acNewDatabaseFormatAccess2007,
                                        Type.Missing);
                    releaseObject(AccessAPP);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Could not create: "+fileName);
                }
                             
            }
            public static bool assertAccessDB()
            {
                try
                {
                    //asserts the file exists
                    if (File.Exists(path + "VDI.accdb"))
                    {
                        result = true;
                        Console.WriteLine("Access file saved as: " + fileName);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error " + e.Message);
                }

                try
                {
                    Available.DeleteAllFilesInPath(path);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error " + e.Message);

                }
                return result;
            }

        }
    }
}
