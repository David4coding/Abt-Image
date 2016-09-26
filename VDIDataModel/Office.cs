using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using Microsoft.Office.Core;
using System.DirectoryServices;
using Access = Microsoft.Office.Interop.Access;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Reflection;

namespace VDIDataModel
{
    public static class Office
    {
        //path for the test to be written
        private static string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                                 + @"\VDI_FRAMEWORK_RUNNER\";

        /// <summary>
        /// releases  the object from memory and calls the GC
        /// </summary>
        /// <param name="obj">Object to be released</param>
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
            public static bool CreateWord()
            {
                object oMissing = System.Reflection.Missing.Value;
                object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
                bool result = false;
                string filename = @"VDI.docx";


                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                try
                {

                    //Start Word and create a new document.
                    Word._Application oWord;
                    Word._Document oDoc;
                    oWord = new Word.Application();
                    oWord.Visible = false;
                    oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing);

                    //Insert a paragraph at the beginning of the document.
                    Word.Paragraph oPara1;
                    oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
                    oPara1.Range.Text = "VDI MICROSOFT TEST";
                    oPara1.Range.Font.Bold = 1;
                    oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
                    oPara1.Range.InsertParagraphAfter();

                    oDoc.SaveAs(path + filename);

                    oDoc.Close();
                    oWord.Quit();
                    //realease from memory and call the GC
                    releaseObject(oDoc);
                    releaseObject(oWord);

                    //asserts the file exists
                    if (File.Exists(path + "VDI.docx"))
                    {
                        result = true;
                        Console.WriteLine("Word file saved as: " + filename);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Word Doc could not be created at this time: " + e.Message);

                }

                return result;
            }
        }
        public static class ExcelWrapper
        {
            public static bool CreateExcel()
            {
                object oMissing = System.Reflection.Missing.Value;
                object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */
                bool result = false;
                string fileName = @"VDI.xlsx";


                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                try
                {
                    Excel.Application ExcelApp = new Excel.Application();

                    if (ExcelApp == null)
                    {
                        Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                        return false;
                    }
                    ExcelApp.Visible = false;

                    Excel.Workbook excelWorkbook = ExcelApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets[1];

                    if (excelWorksheet == null)
                    {
                        Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
                    }

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

                    excelWorkbook.SaveAs(path + fileName, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing,
                        Type.Missing, true, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

                    //asserts the file exists
                    if (File.Exists(path + "VDI.xlsx"))
                    {
                        result = true;
                        Console.WriteLine("Excel file saved as: " + fileName);
                    }

                    excelWorkbook.Close();// change missValue to null
                    ExcelApp.Quit();

                    releaseObject(excelWorksheet);
                    releaseObject(excelWorkbook);
                    releaseObject(ExcelApp);

                }
                catch (Exception e)
                {
                    Console.WriteLine(" Opening Excel, \n Add new content to Excel, \n Saving Failed: " + fileName + "\n" + e.Message); // Success

                }

                return result;
            }


        }
        public static class PowerPointWrapper
        {
            public static bool CreatePowerPoint()
            {
                bool result = false;
                var fileName = "VDI.pptx";
                try
                {

                    PowerPoint.Application ppApp = new PowerPoint.Application();
  
                    PowerPoint.Presentations ppPresens = ppApp.Presentations;
                    PowerPoint.Presentation objPres = ppPresens.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
                    PowerPoint.Slides objSlides = objPres.Slides;
                   
                    PowerPoint.Slide objSlide = objSlides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
                    PowerPoint.TextRange objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
                    objTextRng.Text = "VDI POWER POINT TEST";
                    objTextRng.Font.Name = "Comic Sans MS";
                    objTextRng.Font.Size = 48;

                    objPres.SaveAs(path + fileName, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);

                    objPres.Close();
                    ppApp.Quit();

                    //asserts the file exists
                    if (File.Exists(path + "VDI.pptx"))
                    {
                        result = true;
                        Console.WriteLine("PowerPoint file saved as: " + fileName);
                    }

                    releaseObject(objTextRng);
                    releaseObject(objSlide);
                    releaseObject(objSlides);
                    releaseObject(objPres);
                    releaseObject(ppPresens);
                    releaseObject(ppApp);
                    
                }
                catch (Exception e)
                {
                    Console.WriteLine("Creating PowerPoint Failed: " + fileName + "\n" + e.Message);

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
                    //AccessAPP.NewCurrentDatabase(database,
                    //                       Access.AcNewDatabaseFormat.acNewDatabaseFormatUserDefault,
                    //                      Type.Missing);
                    AccessAPP.NewCurrentDatabase(path+fileName,
                                         Access.AcNewDatabaseFormat.acNewDatabaseFormatAccess2007,
                                        Type.Missing);
                    //asserts the file exists
                    if (File.Exists(path + "VDI.accdb"))
                    {
                        result = true;
                        Console.WriteLine("access file saved as: " + fileName);
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine("could not create access db: " + e.Message);
                }

                return result;
              
            }
        }
    }
}
