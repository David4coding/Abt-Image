using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.DirectoryServices;
using Access = Microsoft.Office.Interop.Access;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Reflection;
using Microsoft.Office.Core;

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
                Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        public static class OutlookWrapper
        {
            private static Outlook.Application outlook;

            public static bool isLoggedIn()
            {
                outlook = new Microsoft.Office.Interop.Outlook.Application();
                String OutlookUser = outlook.Application.Session.CurrentUser.Name;
                DirectoryEntry de = new DirectoryEntry("WinNT://" + Environment.UserDomainName + "/" + Environment.UserName);
                string WindowsUser = de.Properties["fullName"].Value.ToString();

                if (OutlookUser == WindowsUser)
                {
                    Console.WriteLine("Outlook User: "+ WindowsUser);
                    return true;
                }


                return false;
            }

        }
        public static class WordWrapper
        {
            /// <summary>
            /// creates a word file .docx
            /// </summary>
            /// <param name="fileName">template file for the Copy funct</param>
      
            public static Word._Application oWord;
            public static Word._Document oDoc;
            public static Word.Paragraph oPara1;

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
                    oWord = new Word.Application();
                    oWord.Visible = false;
                    oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Word Doc could not be created at this time: " + e.Message);

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
                    Console.WriteLine("paragraph could not be added at this time: " + e.Message);

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
                    Console.WriteLine("Word Doc could not be created at this time: " + e.Message);

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
                    Console.WriteLine("Word Doc could not be created at this time: " + e.Message);
                }
                return result;
            }
        }
        public static class ExcelWrapper
        {
            public static Excel.Application ExcelApp = new Excel.Application();

            public static Excel.Workbook excelWorkbook = ExcelApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            public static Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets[1];

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
                Excel.Range aRange = excelWorksheet.get_Range("C1", "C7");

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
                    excelWorkbook.SaveAs(path + fileName, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing,
                    Type.Missing, true, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

                    excelWorkbook.Close();// change missValue to null
                    ExcelApp.Quit();
                    //relese from memory
                    releaseObject(excelWorksheet);
                    releaseObject(excelWorkbook);
                    releaseObject(ExcelApp);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Excel Word Doc could not be created at this time: " + e.Message);
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
                catch (Exception)
                {

                    throw;
                }
                return result;
            }


        }
        public static class PowerPointWrapper
        {
            public static PowerPoint.Application ppApp = new PowerPoint.Application();
            public static PowerPoint.Presentations ppPresens ;
            public static PowerPoint.Presentation objPres ;
            public static PowerPoint.Slides objSlides ;
            public static PowerPoint.Slide objSlide ;
            public static PowerPoint.TextRange objTextRng;

            public static bool result = false;
            public static string fileName = "VDI.pptx";

            public static bool CreatePowerPoint()
            {
             
                try
                {   //create powerpoint
                    ppPresens = ppApp.Presentations;
                    objPres = ppPresens.Add(MsoTriState.msoTrue);
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
                    objSlide = objSlides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
                    objTextRng = objSlide.Shapes[1].TextFrame.TextRange;

                    objTextRng.Text = "VDI POWER POINT TEST";
                    objTextRng.Font.Name = "Comic Sans MS";
                    objTextRng.Font.Size = 48;
                }
                catch (Exception e)
                {
                    Console.WriteLine("Adding Text o the Slide Failed: " + fileName + "\n" + e.Message);

                }
            }

            public static void savePowerPoint()
            {
                try
                {
                    //saving powerpoint
                    objPres.SaveAs(path + fileName, PowerPoint.PpSaveAsFileType.ppSaveAsDefault,
                        MsoTriState.msoTrue);
                    
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
                    Console.WriteLine("Save PowerPoint Failed at this time: " + fileName + "\n" + e.Message);

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
            public static bool CreateAccess()
            {
                bool result = false;
                var fileName = "VDI.accdb";
                try
                {
                    Access.Application AccessAPP = new Access.Application();
                    AccessAPP.NewCurrentDatabase(path+fileName,
                                         Access.AcNewDatabaseFormat.acNewDatabaseFormatAccess2007,
                                        Type.Missing);
                    //asserts the file exists
                    if (File.Exists(path + "VDI.accdb"))
                    {
                        result = true;
                        Console.WriteLine("Access file saved as: " + fileName);
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine("Could not create access db: " + e.Message);
                    Available.DeleteAllFilesInPath(path);
                }

                return result;
              
            }
        }
    }
}
