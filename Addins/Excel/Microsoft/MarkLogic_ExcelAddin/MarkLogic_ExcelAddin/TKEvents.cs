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
        public void app_SheetChange(object Sh, Excel.Range Target)
        {
            object missing = Type.Missing;
            Excel.Worksheet sheet = (Excel.Worksheet)Sh;

            string changedRange = Target.get_Address(missing, missing,
                Excel.XlReferenceStyle.xlA1, missing, missing);

            MessageBox.Show("The value of " + sheet.Name + ":" +
                changedRange + " was changed.");
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
                //MessageBox.Show("ERROR: " + sheetName);
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
