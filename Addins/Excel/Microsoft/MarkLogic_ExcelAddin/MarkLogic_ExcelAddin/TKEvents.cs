using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System;

namespace MarkLogic_ExcelAddin
{
    public partial class UserControl1 : UserControl
    {
        //BEGIN ADD/REMOVE EVENTS
        public string addChartObjectMouseDownEvents(string sheetName)
        {
            string message = "";
            object missing = Type.Missing;
            try
            {
                Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
                Excel.ChartObjects cs = (Excel.ChartObjects)ws.ChartObjects(missing);

                foreach (Excel.ChartObject c in cs)
                {
                    addEmbeddedChartEvents(c.Chart);
                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                //MessageBox.Show("IN THE ADDDIN ERROR");
            }

            return message;
        }

        public void addEmbeddedChartEvents(Excel.Chart chart)
        {
            chart.MouseDown -= new Microsoft.Office.Interop.Excel.ChartEvents_MouseDownEventHandler(chart_MouseDown);
            chart.MouseDown += new Microsoft.Office.Interop.Excel.ChartEvents_MouseDownEventHandler(chart_MouseDown);  
        }

        public string removeChartObjectMouseDownEvents(string sheetName)
        {
            string message = "";
            object missing = Type.Missing;
            try
            {
                Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
                Excel.ChartObjects cs = (Excel.ChartObjects)ws.ChartObjects(missing);

                foreach (Excel.ChartObject c in cs)
                {
                    removeEmbeddedChartEvents(c.Chart);
                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                //MessageBox.Show("IN THE ADDDIN ERROR");
            }

            return message;
        }

        public void removeEmbeddedChartEvents(Excel.Chart chart)
        {
            chart.MouseDown -= new Microsoft.Office.Interop.Excel.ChartEvents_MouseDownEventHandler(chart_MouseDown);
        }

        
        //END ADD/REMOVE EVENTS

        //BEGIN EVENT HANDLERS
        void app_WorkbookActivate(Excel.Workbook wb)
        {
            string workbookName = "";
            try
            {
                workbookName = wb.Name;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                workbookName = "error: " + errorMsg;
            }

            notifyWorkbookActivate(workbookName);


        }

        void app_WorkbookAfterXmlExport(Microsoft.Office.Interop.Excel.Workbook wb, Microsoft.Office.Interop.Excel.XmlMap map, string url, Microsoft.Office.Interop.Excel.XlXmlExportResult result)
        {
            string workbookName = "";
            string mapName = map.Name;
           
            try
            {
                workbookName = wb.Name;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                workbookName = "error: " + errorMsg;
            }

            notifyWorkbookAfterXmlExport(workbookName, mapName, url);
        }

        void app_WorkbookAfterXmlImport(Microsoft.Office.Interop.Excel.Workbook wb, Microsoft.Office.Interop.Excel.XmlMap map, bool isRefresh, Microsoft.Office.Interop.Excel.XlXmlImportResult Result)
        {
            string workbookName = "";
            string mapName = map.Name;
            string refresh="false";

            
            try
            {
                workbookName = wb.Name;
                if (isRefresh)
                {
                    refresh = "true";
                }

            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                workbookName = "error: " + errorMsg;
            }

            notifyWorkbookAfterXmlImport(workbookName, mapName, refresh);
        }

        void app_WorkbookBeforeClose(Microsoft.Office.Interop.Excel.Workbook wb, ref bool Cancel)
        {
            string workbookName = "";
            try
            {
                workbookName = wb.Name;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                workbookName = "error: " + errorMsg;
            }

            notifyWorkbookBeforeClose(workbookName);

        }

        void app_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook wb, bool SaveAsUI, ref bool Cancel)
        {
            string workbookName = "";
            try
            {
                workbookName = wb.Name;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                workbookName = "error: " + errorMsg;
            }

            notifyWorkbookBeforeSave(workbookName);

        }

        void app_WorkbookBeforeXmlExport(Microsoft.Office.Interop.Excel.Workbook wb, Microsoft.Office.Interop.Excel.XmlMap map, string url, ref bool cancel)
        {
            string workbookName = "";
            string mapName = map.Name;
           
            try
            {
                workbookName = wb.Name;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                workbookName = "error: " + errorMsg;
            }

            notifyWorkbookBeforeXmlExport(workbookName, mapName, url);
        }

        void app_WorkbookBeforeXmlImport(Microsoft.Office.Interop.Excel.Workbook wb, Microsoft.Office.Interop.Excel.XmlMap map, string url, bool isRefresh, ref bool Cancel)
        {
            string workbookName = "";
            string mapName = map.Name;
            string refresh = "false";


            try
            {
                workbookName = wb.Name;
                if (isRefresh)
                {
                    refresh = "true";
                }

            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                workbookName = "error: " + errorMsg;
            }

            notifyWorkbookBeforeXmlImport(workbookName, mapName, refresh);
        }

        void app_WorkbookDeactivate(Microsoft.Office.Interop.Excel.Workbook wb)
        {
            string workbookName = "";
            try
            {
                workbookName = wb.Name;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                workbookName = "error: " + errorMsg;
            }

            notifyWorkbookDeactivate(workbookName);

        }

        void app_WorkbookNewSheet(Microsoft.Office.Interop.Excel.Workbook wb, object Sh)
        {
            string workbookName = "";
            string sheetName = "";
            try
            {
                workbookName = wb.Name;
                Excel.Worksheet ws = (Excel.Worksheet)Sh;
                sheetName = ws.Name;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                workbookName = "error: " + errorMsg;
            }

            notifyWorkbookNewSheet(workbookName, sheetName);

        }

        public void app_WorkbookOpen(Excel.Workbook wb)
        {
            string workbookName = "";
            try
            {
                workbookName = wb.Name;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                workbookName = "error: " + errorMsg;
            }
            //MessageBox.Show("notifying");
            //tricky.  this will fire, but if the page has'nt loaded yet, the event wont' get caught
            //instead of using workbook open, for initial of app, onload event 
            //should check sheet type and initialize the first sheet with chart events

            notifyWorkbookOpen(workbookName);

        }



        public void app_SheetActivate(object Sh)
        {
            //check for chart sheet (that can use this function)
            //BREAK THIS UP IN JS FOR MORE FLEXIBILITY
            //check for embedded charts (have to explicitly add event handler for activate)
            //let event fire for worksheet/chartsheet as required
            //subsequent embedded chart activation will be handled by other function

            //MessageBox.Show("In Sheet Activate");

            string sheetName = "";
            try
            {
                Excel.Workbook awb = Globals.ThisAddIn.Application.ActiveWorkbook;
               
                if (awb.ActiveSheet is Excel.Worksheet)
                {
                    Excel.Worksheet ws = (Excel.Worksheet)Sh;
                    sheetName = ws.Name;
                }
                else if (awb.ActiveSheet is Excel.Chart)  //wont work casting SH to worksheet chart in is
                {
                    Excel.Chart chart = (Excel.Chart)Sh;
                    sheetName = chart.Name;
                }
                else
                {
                    //MessageBox.Show("what the hell is it?");
                    //not able to determine name
                    //some meaningful message here?
                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                sheetName = "error: " + errorMsg;
            }

            notifySheetActivate(sheetName);
        }

        void app_SheetBeforeDoubleClick(object Sh, Microsoft.Office.Interop.Excel.Range Target, ref bool Cancel)
        {
            string sheetName="";
            string range = "";
            try
            {
                object missing = Type.Missing;
                Excel.Worksheet sheet = (Excel.Worksheet)Sh;
                sheetName=sheet.Name;
                range = Target.get_Address(missing, missing, Excel.XlReferenceStyle.xlA1, missing, missing);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                sheetName = "error: " + errorMsg;
            }

            notifySheetBeforeDoubleClick(sheetName, range);
        }

        void app_SheetBeforeRightClick(object Sh, Microsoft.Office.Interop.Excel.Range Target, ref bool Cancel)
        {
            string sheetName = "";
            string range = "";
            try
            {
                object missing = Type.Missing;
                Excel.Worksheet sheet = (Excel.Worksheet)Sh;
                sheetName = sheet.Name;
                range = Target.get_Address(missing, missing, Excel.XlReferenceStyle.xlA1, missing, missing);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                sheetName = "error: " + errorMsg;
            }

            notifySheetBeforeRightClick(sheetName, range);
        }

        void app_SheetDeactivate(object Sh)
        {
             //MessageBox.Show("In Sheet Deactivate");
             string sheetName = "";
             try{

                 try
                 {

                     if ((Excel.Worksheet)Sh is Excel.Worksheet)
                     {
                         Excel.Worksheet ws = (Excel.Worksheet)Sh;
                         sheetName = ws.Name;
                     }
                 }
                 catch (Exception e)
                 {
                     string donothing_removewarning = e.Message;
                     try
                     {
                         if ((Excel.Chart)Sh is Excel.Chart)  //wont work casting SH to worksheet chart in is
                         {
                             Excel.Chart chart = (Excel.Chart)Sh;
                             sheetName = chart.Name;
                         }
                     }
                     catch (Exception e2)
                     {
                         string donothing_removewarning2 = e2.Message;
                     }
                 }
                
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                sheetName = "error: " + errorMsg;
            }

            notifySheetDeactivate(sheetName);
        }

        public void app_SheetChange(object Sh, Excel.Range Target)
        {
            string range = "";
            try
            {
                object missing = Type.Missing;
                Excel.Worksheet sheet = (Excel.Worksheet)Sh;

                range = Target.get_Address(missing, missing, Excel.XlReferenceStyle.xlA1, missing, missing);

                MessageBox.Show("The value of " + sheet.Name + ":" +
                    range + " was changed.");
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                range = "error: " + errorMsg;
            }

            notifySheetChange(range);
        }

        void app_SheetSelectionChange(object Sh, Microsoft.Office.Interop.Excel.Range Target)
        {
            //MessageBox.Show("In Sheet Selection Change");
            string sheetName = "";
            string rangeName = "";
            try
            {
                    if ((Excel.Worksheet)Sh is Excel.Worksheet)
                    {
                        Excel.Worksheet ws = (Excel.Worksheet)Sh;
                        sheetName = ws.Name;
                        Excel.Name nm = (Excel.Name)Target.Name;
                        rangeName = nm.Name;
                    
                        notifyRangeSelected(rangeName);
                    }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                string donothing_removewarning = "error: " + errorMsg;
                //will error when a named range is not selected as name does not exist for other/individual cells
                //MessageBox.Show("ERROR: " + sheetName+ " "+donothing_removewarning);
            }

            //notifyRangeSelected(rangeName);
            
        }

        public void chart_MouseDown(int Button, int Shift, int x, int y)
        {
            string message = "";
            try
            {
                notifyChartObjectMouseDown(Globals.ThisAddIn.Application.ActiveChart.Name);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                //see what i do in powerpoint wrt events
                //MessageBox.Show("ERROR " + e.Message);
            }
        }

        /*
        public void chart_Activate()
        {
            MessageBox.Show("CHART NAME: " + Globals.ThisAddIn.Application.ActiveWorkbook.ActiveChart.Name);
        }

        void chart_SelectEvent(int ElementID, int Arg1, int Arg2)
        {
            MessageBox.Show("IN SELECT EVENT");
            if (Excel.XlChartItem.xlAxis == (Excel.XlChartItem)ElementID)
            {
                if (Excel.XlAxisGroup.xlPrimary == (Excel.XlAxisGroup)Arg1)
                {
                    MessageBox.Show("The primary axis of the chart was selected.");
                }
            }
        }
         * */
        //END EVENT HANDLERS

        //BEGIN NOTIFY
        public void notifyWorkbookActivate(string workbookName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("workbookActivate", new String[] { workbookName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("workbookActivateJS: " + res);
                }
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void  notifyWorkbookAfterXmlExport(string workbookName, string mapName, string url)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("workbookAfterXmlExport", new String[] { workbookName, mapName, url });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("workbookAfterXmlExportJS: " + res);
                }
                // MessageBox.Show("CALLED THE JS");
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifyWorkbookAfterXmlImport(string workbookName, string mapName, string refresh)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("workbookAfterXmlImport", new String[] { workbookName, mapName, refresh});
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("workbookAfterXmlImportJS: " + res);
                }
                // MessageBox.Show("CALLED THE JS");
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifyWorkbookBeforeClose(string workbookName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("workbookBeforeClose", new String[] { workbookName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("workbookBeforeCloseJS: " + res);
                }
                // MessageBox.Show("CALLED THE JS");
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifyWorkbookBeforeSave(string workbookName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("workbookBeforeSave", new String[] { workbookName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("workbookBeforeSaveJS: " + res);
                }
                // MessageBox.Show("CALLED THE JS");
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifyWorkbookBeforeXmlExport(string workbookName, string mapName, string url)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("workbookBeforeXmlExport", new String[] { workbookName, mapName, url });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("workbookBeforeXmlExportJS: " + res);
                }
                // MessageBox.Show("CALLED THE JS");
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifyWorkbookBeforeXmlImport(string workbookName, string mapName, string url)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("workbookBeforeXmlImport", new String[] { workbookName, mapName, url });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("workbookBeforeXmlImportJS: " + res);
                }
                // MessageBox.Show("CALLED THE JS");
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifyWorkbookDeactivate(string workbookName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("workbookDeactivate", new String[] { workbookName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("workbookDeactivateJS: " + res);
                }
                // MessageBox.Show("CALLED THE JS");
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifyWorkbookNewSheet(string workbookName, string sheetName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("workbookNewSheet", new String[] { workbookName, sheetName});
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("workbookNewSheetJS: " + res);
                }
                // MessageBox.Show("CALLED THE JS");
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifyWorkbookOpen(string workbookName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("workbookOpen", new String[] { workbookName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("workbookOpenJS: " + res);
                }
                // MessageBox.Show("CALLED THE JS");
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifySheetActivate(string sheetName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("sheetActivate", new String[] { sheetName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("sheetActivateJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifySheetBeforeDoubleClick(string sheetName, string range)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("sheetBeforeDoubleClick", new String[] { sheetName, range });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("sheetBeforeDoubleClickJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifySheetBeforeRightClick(string sheetName, string range)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("sheetBeforeRightClick", new String[] { sheetName, range });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("sheetBeforeRightClickJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifySheetChange(string rangeName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("sheetChange", new String[] { rangeName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("sheetChangeJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifySheetDeactivate(string sheetName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("sheetDeactivate", new String[] { sheetName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("sheetDeactivateJS: " + res);
                }

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show(donothing_removewarning);
            }
        }

        public void notifyChartObjectMouseDown(string chartName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("chartObjectMouseDown", new String[] { chartName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("chartObjectMouseDownJS: " + res);
                }
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        public void notifyRangeSelected(string rangeName)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("rangeSelected", new String[] { rangeName });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show("rangeSelectedJS: " + res);
                }
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
        }

        //END NOTIFY
    }
}
