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
 * Program.cs - A simple app to start/end Excel used for testing .js api
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
                object file=args[0];
                //For Save As
                //object file = @"c:\unitTestAddin\outputs\test.xlsx";             

                Excel.Application excelApp;
                Excel.Workbook wb;
                
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                wb = excelApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;
                
                
                wb.SaveAs(file, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);

                System.Threading.Thread.Sleep(10000);

                wb.Save();
                excelApp.Workbooks.Close();
                excelApp.Quit();

            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("ERROR" + e.Message);

            }
        }
    }
}
