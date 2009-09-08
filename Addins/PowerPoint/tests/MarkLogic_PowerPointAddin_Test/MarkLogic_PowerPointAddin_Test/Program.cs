/*Copyright 2009 Mark Logic Corporation

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
 * 
 * Program.cs - A simple app to start/quit PowerPoint 
 *              used for testing .xqy and .js APIs
*/
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.IO;
using System;
using Office = Microsoft.Office.Core;
using Microsoft.Win32;
using PPT = Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
using System.Linq;
using System.Text;



namespace MarkLogic_PowerPointAddin_Test
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\unitTestAddin\pptx";
            string title = "VALIDATE_THREE.pptx";
            string message = "";
            object missing = Type.Missing;
            string tmpdoc = "";

            try
            {
                tmpdoc = path + "\\"+title;
                PPT.Application pptApp;
                PPT.Presentation ppt;
                //Excel.Application excelApp;
                //Excel.Workbook wb;

                pptApp = new PPT.Application();
                pptApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue ;
              
                //downloadFile(url, tmpdoc, user, pwd);
                ppt = pptApp.Presentations.Open(tmpdoc, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue);
                System.Threading.Thread.Sleep(20000);
                pptApp.Quit();

            }
            catch (Exception e)
            {
                //not always true, need to improve error handling or message or both
                string origmsg = "A presentation with the name '" + title + "' is already open. You cannot open two documents with the same name, even if the documents are in different \nfolders. To open the second document, either close the document that's currently open, or rename one of the documents.";
                MessageBox.Show(origmsg);
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            

         //   return message;

            /*
            try
            {

                object install = true;
                object missing = System.Reflection.Missing.Value;

                //For Save As
                object file = @"c:\unitTestAddin\outputs\test.pptx";
                PPT.Application pptApp;
                PPT.Presentation pres;
                //Excel.Application excelApp;
                //Excel.Workbook wb;

                pptApp = new PPT.Application();
                pptApp.Visible = true;
                pres = pptApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
                PPT.Presentation p = 

                excelApp = new Excel.Application();
                excelApp.Visible = true;
                wb = excelApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;

                wb.SaveAs(file, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);

                System.Threading.Thread.Sleep(20000);

                wb.Save();
                excelApp.Workbooks.Close();
                excelApp.Quit();

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("ERROR" + e.Message);

            }
             * */
        }
    }
}
