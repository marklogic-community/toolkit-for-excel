using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.IO;
using System;
using Office = Microsoft.Office.Core;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;


namespace MarkLogic_ExcelAddin_Test
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
             
                object install = true;
                object missing = System.Reflection.Missing.Value;
                
                //For Save As
                object file = @"c:\unitTestAddin\outputs\test.xlsx";
                Excel.Application excelApp;
                Excel.Workbook wb;
                
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                wb = excelApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
              // MessageBox.Show("BOOKNAME: "+wb.Name);
                Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;
              // MessageBox.Show("SHEETNAME: "+ws.Name);
                
                wb.SaveAs(file, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);

            
                // wordApp.Application.AddIns.Add(@"C:\Program Files\MarkLogic\MarkLogic_WordAddin\MarkLogic_WordAddin.dll",ref install);
                //object testapp = "MarkLogic_WordAddin";

                //wordDoc.SaveAs(ref file, ref FileFormat, ref LockComments, ref pwd, ref addtorecentfiles, ref writepswd, ref readonlyrecommend, ref embedtruetypefont, ref savenativepicformat, ref saveformsdata, ref saveasaocelletter, ref encoding, ref insertlinebreaks, ref allowsus, ref lineend, ref addbidi);

                //Timer here?  //require time for page to load before saving
                //thread sleep is in milliseconds (1000 = 1 sec)
                //below set for 20 secs
                System.Threading.Thread.Sleep(20000);
                

                //wb.Close(true, false, false);
                wb.Save();
                excelApp.Workbooks.Close();
                //wb.Close(true, false, false);
                excelApp.Quit();

              /*  object FileFormat = Word.WdSaveFormat.wdFormatDocumentDefault;
                object LockComments = false;
                object pwd = "";
                object addtorecentfiles = true;
                object writepswd = "";
                object readonlyrecommend = false;
                object embedtruetypefont = false;
                object savenativepicformat = false;
                object saveformsdata = false;
                object saveasaocelletter = false;
                object encoding = false;
                object insertlinebreaks = false;
                object allowsus = false;
                object lineend = false;
                object addbidi = false;

                //For Close
                object saveChanges = 
                    Word.WdSaveOptions.wdSaveChanges;
                object originalFormat = Word.WdOriginalFormat.wdOriginalDocumentFormat;
                object routeDocument = true;


                //  Start Word and create a new document.
                Word.Application wordApp;
                Word.Document wordDoc;
                wordApp = new Word.Application();
                wordApp.Visible = true;
                wordDoc = wordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                // wordApp.Application.AddIns.Add(@"C:\Program Files\MarkLogic\MarkLogic_WordAddin\MarkLogic_WordAddin.dll",ref install);
                object testapp = "MarkLogic_WordAddin";

                wordDoc.SaveAs(ref file, ref FileFormat, ref LockComments, ref pwd, ref addtorecentfiles, ref writepswd, ref readonlyrecommend, ref embedtruetypefont, ref savenativepicformat, ref saveformsdata, ref saveasaocelletter, ref encoding, ref insertlinebreaks, ref allowsus, ref lineend, ref addbidi);

                //Timer here?  //require time for page to load before saving
                //thread sleep is in milliseconds (1000 = 1 sec)
                //below set for 20 secs
                System.Threading.Thread.Sleep(20000);
                // MessageBox.Show("TEST");

                //wordDoc.Close(ref saveChanges, ref originalFormat, ref routeDocument);
                wordApp.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
*/

                //              Microsoft.Office.Core.COMAddIn comAddIn= null;
                //              Microsoft.Office.Core.COMAddIns addins;
                //              addins = wordApp.COMAddIns;

                /*              foreach (Microsoft.Office.Core.COMAddIn a in addins)
                                {
                                    //MessageBox.Show("" + a.Description + " " + a.Guid + " " + a.ProgId + " ");
                                    if (a.ProgId.Equals("MarkLogic_WordAddin"))
                                        comAddIn = a;
                  
                                }

                                comAddIn.Connect = true;  
                */


                /*              foreach(Word.AddIn ad in wordApp.AddIns)
                                {  // if (ad.Installed == true)
                                   // {
                                    // shows .odt only
                                        MessageBox.Show(ad.Name + "IS INSTALLED");
                                   // }
                                }
                */

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("ERROR" + e.Message);

            }
        }
    }
}
