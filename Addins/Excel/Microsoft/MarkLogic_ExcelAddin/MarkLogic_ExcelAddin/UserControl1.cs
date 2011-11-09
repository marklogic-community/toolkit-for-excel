/*Copyright 2009-2011 Mark Logic Corporation

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
 * UserControl1.cs - the api called from MarkLogicExcelAddin.js.  The methods here map directly to functions in the .js.
 * 
*/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using Excel = Microsoft.Office.Interop.Excel;
using VBIDE = Microsoft.Vbe.Interop;

using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.IO;

//using DocumentFormat.OpenXml.Packaging; //OpenXML sdk
using Office = Microsoft.Office.Core;
using Microsoft.Win32;
using Tools = Microsoft.Office.Tools.Excel;


namespace MarkLogic_ExcelAddin
{
    [ComVisible(true)]
  
        public partial class UserControl1 : UserControl
        {

            private AddinConfiguration ac = AddinConfiguration.GetInstance();
            private string webUrl = "";
            private bool debug = false;
            private bool debugMsg = false;
            private string color = "";
            //private string addinVersion = "@MAJOR_VERSION.@MINOR_VERSION@PATCH_VERSION";  
            private string addinVersion = "2.0-1"; 
            HtmlDocument htmlDoc;

            /*private const int CF_ENHMETAFILE = 14;
              private const int CF_METAFILEPICT = 3;
              [DllImport("user32.dll")]
              private static extern bool OpenClipboard(IntPtr hWndNewOwner);
              [DllImport("user32.dll")]
              private static extern int IsClipboardFormatAvailable(int wFormat);
              [DllImport("user32.dll")]
              private static extern IntPtr GetClipboardData(int wFormat);
              [DllImport("user32.dll")]
              private static extern int CloseClipboard();
              [DllImport("gdi32.dll")]
              static extern IntPtr CopyEnhMetaFile(IntPtr hemfSrc, IntPtr hNULL);
              [System.Runtime.InteropServices.DllImport("gdi32")]
              public static extern int GetEnhMetaFileBits(int hemf, int cbBuffer, byte[] lpbBuffer);
            **/
            //static extern IntPtr CopyEnhMetaFile(IntPtr hemfSrc, string lpszFile);

            public UserControl1()
            {
                InitializeComponent();
                //bool regEntryExists = checkUrlInRegistry();
                webUrl = ac.getWebURL();

                if (webUrl.Equals(""))
                {
                    MessageBox.Show("                                   Unable to find configuration info. \n\r " +
                                    " Please see the README for how to add configuration info for your system. \n\r " +
                                    "           If problems persist, please contact your system administrator.");
                }
                else
                {
                    color = TryGetColorScheme().ToString();
                    webBrowser1.AllowWebBrowserDrop = false;
                    webBrowser1.IsWebBrowserContextMenuEnabled = false;
                    webBrowser1.WebBrowserShortcutsEnabled = false;
                    webBrowser1.ObjectForScripting = this;
                    webBrowser1.Navigate(webUrl);
                    webBrowser1.ScriptErrorsSuppressed = true;

                    this.webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser1_DocumentCompleted);

                    if (ac.getEventsEnabled())
                    { 
                        Excel.Application app = Globals.ThisAddIn.Application;

                        app.WorkbookActivate += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookActivateEventHandler(app_WorkbookActivate);
                        app.WorkbookAfterXmlExport += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookAfterXmlExportEventHandler(app_WorkbookAfterXmlExport);
                        app.WorkbookAfterXmlImport += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookAfterXmlImportEventHandler(app_WorkbookAfterXmlImport);
                        app.WorkbookBeforeClose += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeCloseEventHandler(app_WorkbookBeforeClose);
                        app.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(app_WorkbookBeforeSave);
                        app.WorkbookBeforeXmlExport += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeXmlExportEventHandler(app_WorkbookBeforeXmlExport);
                        app.WorkbookBeforeXmlImport += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeXmlImportEventHandler(app_WorkbookBeforeXmlImport);
                        app.WorkbookDeactivate += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookDeactivateEventHandler(app_WorkbookDeactivate);
                        app.WorkbookNewSheet += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookNewSheetEventHandler(app_WorkbookNewSheet);
                        app.WorkbookOpen += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookOpenEventHandler(app_WorkbookOpen);
                        app.SheetActivate += new Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler(app_SheetActivate);
                        app.SheetBeforeDoubleClick += new Microsoft.Office.Interop.Excel.AppEvents_SheetBeforeDoubleClickEventHandler(app_SheetBeforeDoubleClick);
                        app.SheetBeforeRightClick += new Microsoft.Office.Interop.Excel.AppEvents_SheetBeforeRightClickEventHandler(app_SheetBeforeRightClick);
                        app.SheetChange += new Microsoft.Office.Interop.Excel.AppEvents_SheetChangeEventHandler(app_SheetChange);
                        app.SheetDeactivate += new Microsoft.Office.Interop.Excel.AppEvents_SheetDeactivateEventHandler(app_SheetDeactivate);
                        app.SheetSelectionChange += new Microsoft.Office.Interop.Excel.AppEvents_SheetSelectionChangeEventHandler(app_SheetSelectionChange);
                      
                        
                    }

                }
            }

            private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
            {
                if (webBrowser1.Document != null)
                {
                    htmlDoc = webBrowser1.Document;

                    htmlDoc.Click += htmlDoc_Click;

                }
            }

            private void htmlDoc_Click(object sender, HtmlElementEventArgs e)
            {
                if (!(webBrowser1.Parent.Focused))
                {
                    webBrowser1.Parent.Focus();
                    webBrowser1.Document.Focus();
                }
                
            }

            private bool checkUrlInRegistry()
            {
                RegistryKey regKey1 = Registry.CurrentUser;
                regKey1 = regKey1.OpenSubKey(@"MarkLogicAddinConfiguration\Word");
                bool keyExists = false;
                if (regKey1 == null)
                {
                    if (debugMsg)
                        MessageBox.Show("KEY IS NULL");

                }
                else
                {
                    if (debugMsg)
                        MessageBox.Show("KEY IS: " + regKey1.GetValue("URL"));

                    webUrl = (string)regKey1.GetValue("URL");
                    if (!((webUrl.Equals("")) || (webUrl == null)))
                        keyExists = true;
                }
                return keyExists;
            }

            public enum ColorScheme : int
            {
                Blue = 1,
                Silver = 2,
                Black = 3,
                Unknown = 4
            };

            public ColorScheme TryGetColorScheme()
            {
                //assume default - theme registry key not always set on install of Office
                //set for sure once user sets color scheme manually from button
                ColorScheme CurrentColorScheme = (ColorScheme)Enum.Parse(typeof(ColorScheme), "1");
                try
                {
                    Microsoft.Win32.RegistryKey rootKey = Microsoft.Win32.Registry.CurrentUser;
                    Microsoft.Win32.RegistryKey registryKey = rootKey.OpenSubKey("Software\\Microsoft\\Office\\12.0\\Common");
                    if (registryKey == null) return ColorScheme.Unknown;

                    CurrentColorScheme =
                        (ColorScheme)Enum.Parse(typeof(ColorScheme), registryKey.GetValue("Theme").ToString());
                }
                catch
                { }

                return CurrentColorScheme;
            }

            public String getOfficeColor()
            {
                return color;
            }

            public String getAddinVersion()
            {
                return addinVersion;
            }

            public String getBrowserUrl()
            {
                return webUrl;
            }

            public String getCustomXMLPartIds()
            {

                string ids = "";

                try
                {
                    Excel.Workbook wkbk = Globals.ThisAddIn.Application.ActiveWorkbook;
                    int count = wkbk.CustomXMLParts.Count;

                    foreach (Office.CustomXMLPart c in wkbk.CustomXMLParts)
                    {
                        if (c.BuiltIn.Equals(false))
                        {
                            ids += c.Id + " ";

                        }
                    }

                    char[] space = { ' ' };
                    ids = ids.TrimEnd(space);
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    ids = "error: " + errorMsg;
                }

                if (debug)
                    ids = "error";

                return ids;
            }


            public String getCustomXMLPart(string id)
            {

                string custompiecexml = "";

                try
                {
                    Excel.Workbook wkbk = Globals.ThisAddIn.Application.ActiveWorkbook;
                    Office.CustomXMLPart cx = wkbk.CustomXMLParts.SelectByID(id);

                    if (cx != null)
                        custompiecexml = cx.XML;

                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    custompiecexml = "error: " + errorMsg;
                }

                if (debug)
                    custompiecexml = "error";

                return custompiecexml;

            }

            public String addCustomXMLPart(string custompiecexml)
            {
                string newid = "";
                try
                {
                    Excel.Workbook wkbk = Globals.ThisAddIn.Application.ActiveWorkbook;
                    Office.CustomXMLPart cx = wkbk.CustomXMLParts.Add(String.Empty, new Office.CustomXMLSchemaCollectionClass());
                    cx.LoadXML(custompiecexml);
                    newid = cx.Id;
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    newid = "error: " + errorMsg;
                }
                if (debug)
                    newid = "error";
                return newid;

            }

            public String deleteCustomXMLPart(string id)
            {
                string message = "";
                try
                {
                    Excel.Workbook wkbk = Globals.ThisAddIn.Application.ActiveWorkbook;
                    foreach (Office.CustomXMLPart c in wkbk.CustomXMLParts)
                    {
                        if (c.BuiltIn.Equals(false) && c.Id.Equals(id))
                        {
                            c.Delete();
                        }

                    }
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                if (debug)
                    message = "error";

                return message;

            }

            public String getActiveWorkbookName()
            {
                string wbname = "";

                try
                {
                    Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                    wbname = wb.Name;
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    wbname = "error: " + errorMsg;
                }

                return wbname;
            }

            public String getActiveWorksheetName()
            {
                string wsname = "";
                try
                {
                    if (Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet is Excel.Worksheet)
                    {
                        Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                        wsname = ws.Name;
                    }
                    else if (Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet is Excel.Chart)
                    {
                        Excel.Chart cs = (Excel.Chart)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                        wsname = cs.Name;
                    }
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    wsname = "error: " + errorMsg;
                    MessageBox.Show(wsname);
                }

                return wsname;
            }

            public String getAllWorkbookNames()
            {
                string workbooks = "";

                try
                {
                    Excel.Workbooks wbs = Globals.ThisAddIn.Application.Workbooks;
                    foreach (Excel.Workbook w in wbs)
                        workbooks += w.Name + "|";

                    int length = workbooks.Length;
                    workbooks = workbooks.Substring(0, length - 1);
                    return workbooks;
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    workbooks = "error: " + errorMsg;
                }

                return workbooks;

            }


            public String getActiveWorkbookWorkSheetNames()
            {
                string worksheets = "";
                try
                {

                    Excel.Sheets ws = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets;//Globals.ThisAddIn.Application.Worksheets;

                    foreach (Excel.Worksheet n in ws)
                        worksheets += n.Name + "|";

                    int length = worksheets.Length;
                    worksheets = worksheets.Substring(0, length - 1);
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    worksheets = "error: " + errorMsg;
                }

                return worksheets;
            }

            public String addWorkbook(string name) //, string subject,string saveas)
            {
                string wbname = "";
                try
                {
                    Excel.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    wb.Title = name;
                    //this names sheet, not book, can't name book unless we do saveas, want to keep?
                    //have to retest this, docs say you have to say, but I see default of Book2, when adding
                    //thru Excel
                    Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];
                    ws.Name = name;
                    wbname = wb.Name;
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    wbname = "error: " + errorMsg;
                }
                return wbname;
            }

            public String addWorksheet(string name)//#sheets as param?
            {
                string message = "";
                Excel.Worksheet ws = null;
                try
                {
                    object missing = Type.Missing;
                    int count = 1; 
                    ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(missing, Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[/*"Sheet1"*/Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Count], count, missing);

                    ws.Name = name;
                    message = name;
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                    //dont't allow adding of worksheet with default name
                    ws.Delete();
                }

                return message;

            }

            public String setActiveWorkbook(string name)
            {
                string message = "";
                try
                {
                    Globals.ThisAddIn.Application.Workbooks[name].Activate();
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                return message;
            }

            public String setActiveWorksheet(string name)
            {
                string message = "";

                try
                {
                    object missing = Type.Missing;
                    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[name]).Select(missing);

                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                return message;
            }

            public String addNamedRange(string coordinate1, string coordinate2, string rngName, string sheetName)
            {

                string message = "";
                object missing = Type.Missing;
                
                try{
                   Excel.Worksheet ws = null;

                   if (sheetName.Equals("active"))
                   {
                      ws = ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                   }
                   else
                   {
                      ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName]; 
                   }

                   Excel.Range rg = ws.get_Range(coordinate1, coordinate2);
                   Excel.Name nm = ws.Names.Add(rngName, rg, true, missing, missing, missing, missing, missing, missing, missing, missing);

                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                return message;
            }

            public String addAutoFilter(string coordinate1,string coordinate2, string sheetName, string criteria1, string v_operator, string criteria2)
            {
                string message = "";
                object missing = Type.Missing;

                try
                {
                    Excel.Worksheet ws = null;

                    if (sheetName.Equals("active"))
                    {
                        ws = ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                    }
                    else
                    {
                        ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
                    }

                    Excel.Range rg = ws.get_Range(coordinate1, coordinate2);
                    rg.AutoFilter(1, "<>", Excel.XlAutoFilterOperator.xlOr, missing, true);
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                return message;


            }

            public String getWorksheetChartNames(string sheetName)
            {
                string message = "";
                string names = "";
                object missing = Type.Missing;
                
                try
                {
                    Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
                    Excel.ChartObjects cs = (Excel.ChartObjects)ws.ChartObjects(missing);

                    
                    foreach (Excel.ChartObject c in cs)
                    {
                            names += c.Name + ":";
                    }

                   
                    if (!(names.Equals("")))
                    {
                        names = names.Substring(0, names.Length - 1);
                    }

                    message = names;
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                    MessageBox.Show(message);
                }
                return message;
            }
            public String getWorksheetNamedRangeRangeNames(string sheetName)
            {
                string message="";
                string names = "";
                try
                {
                    Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
                    Excel.Names ns = ws.Names;
                
                    foreach (Excel.Name n in ns)
                        names += n.Name + ":";

                    if (!(names.Equals("")))
                    {
                        names = names.Substring(0, names.Length - 1);
                    }

                    message = names;
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                return message;
            }

            //for workbooks
            public String getNamedRangeRangeNames()
            {
                string message = "";
                string names = "";
             
                try
                {
                    Excel.Names    ns = Globals.ThisAddIn.Application.ActiveWorkbook.Names;
                    
                    foreach (Excel.Name n in ns)
                        names += n.Name + ":";

                    names = names.Substring(0, names.Length - 1);

                    message = names;
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                return message;

            }

            public String setActiveRangeByName(string rngName)
            {
                String message = "";
                object missing = Type.Missing;
            try{

                Excel.Sheets ws = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets;//Globals.ThisAddIn.Application.Worksheets;
                Excel.Range r = null;

                //loop thru all sheets til we find range, return first, else, give up
                //names have to be unique
                foreach (Excel.Worksheet n in ws)
                {
                    string wsname = n.Name;
                    setActiveWorksheet(wsname);
                    try
                    {

                        r = n.get_Range(rngName, missing);
                        if (r != null)
                        {
                            r.Activate();
                            break;
                        }
                    }
                    catch
                    {
                        
                        r = null;
                    }
                }
            }
            catch(Exception e)
            {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
            }
                
                return message;
            }

            public String clearNamedRange(string rngName)
            {
                String message = "";
                object missing = Type.Missing;
                try
                {
                    String names = getActiveWorkbookWorkSheetNames();
                    Excel.Range r = null;

                    char x = '|';
                    foreach (String name in names.Split(x))
                    {
                        setActiveWorksheet(name);
                        Excel.Worksheet n = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                        try
                        {

                            r = n.get_Range(rngName, missing);
                            if (r != null)
                            {
                                r.Select();
                                r.Clear();
                                break;
                            }
                        }

                        catch (Exception e)
                        {
                            string do_nothing = e.Message;
                            r = null;
                        }

                    }
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                return message;
            }

            public String clearRange(string startcoord, string endcoord)
            {
                string message = "";
                object missing = Type.Missing;
                try
                {
                    Excel.Worksheet w = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;//Globals.ThisAddIn.Application.Worksheets;
                    Excel.Range r = w.get_Range(startcoord, endcoord);
                    r.Clear();
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                return message;

            }

            public String removeNamedRange(string rngName)
            {
                string message = "";
                object missing = Type.Missing;
                try{

                Excel.Names ns = Globals.ThisAddIn.Application.ActiveWorkbook.Names;
                foreach (Excel.Name nDel in ns)
                {
                    if (nDel.Name.EndsWith(rngName))
                    {
                        //MessageBox.Show("Deleting");
                        nDel.Delete();
                    }
                }
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }


                return message;
            }

            public String getSelectedChartName()
            {
                string message = "";
                try
                {
                    message = Globals.ThisAddIn.Application.ActiveChart.Name;
                }
                catch (Exception e)
                {
                    string donothing_removewarning = e.Message;
                   // MessageBox.Show(donothing_removewarning);
                }
                return message;
            }

            public String getSelectedRangeName()
            {
                string message = "";
                try
                {

                    Excel.Range r = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                    Excel.Name nm = (Excel.Name)r.Name;
                    //MessageBox.Show("IN GETSELECTEDRANGENAME" + nm.Name);
                    message = nm.Name;
                }
                catch (Exception e)
                {
                    string donothing_removewarning = e.Message;
                    
                }
                return message;
            }

            public String getSelectedRangeCoordinates()
            {
                string message = "";
                string firstCellCoordinate = "";
                string lastCellCoordinate = "";
                try
                {
                    Excel.Range r = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                  
                    int start = 1;
                    int end = r.Count;
                    int count = 1;
                    //MessageBox.Show("COUNT" + r.Count);

                    foreach (Excel.Range r2 in r)
                    { 
                        if (count == start)
                        {
                            firstCellCoordinate = r2.get_Address(true, true, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, null, null);
                        }
                        else if (count == end)
                        {
                            lastCellCoordinate = r2.get_Address(true, true, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, null, null);
                        }

                        count++;

                    }

                    message = firstCellCoordinate + ":" + lastCellCoordinate;
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }
                return message;
            }

            public String getSelectedCells()
            {
                //MessageBox.Show("IN FUNCTION");
                object missing = Type.Missing;
                string coordinate = "";
                string col = "";
                string row = "";
                string value2 = "";
                string formula = "";
                string message = "";

                try
                {
                    Excel.Range rng = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                    string cells = "[";

                    foreach (Excel.Range r in rng)
                    {
                        row = r.Row + "";
                        col = r.Column + "";
                        coordinate = r.get_Address(r.Row, r.Column, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, missing, missing);
                        value2 = r.Value2 + "";
                        formula = r.Formula + "";
                        //r.AddComment

                        string cell = "";
                        cell = "{ \"rowIdx\": " + "\"" + row + "\""
                             + ",\"colIdx\": " + "\"" + col + "\""
                             + ",\"coordinate\": " + "\"" + coordinate + "\""
                             + ",\"value2\": " + "\"" + value2 + "\""//r.value2
                             + ",\"formula\": " + "\"" + formula + "\"" // r.Formula
                             + "}";

                        cells += cell + ",";
                    }

                    cells = cells.Substring(0, cells.Length - 1);
                    cells += "]";
                    //MessageBox.Show("message: " + cells);

                    message = cells;

                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                return message;
            }

            public String getActiveCell()
            {
                object missing = Type.Missing;
                string col = "";
                string row = "";
                string r1c1 = "";
                string value2 = "";
                string formula = "";
                string message = "";

                try
                {
                    Excel.Range r = Globals.ThisAddIn.Application.ActiveCell;

                    row = r.Row.ToString();
                    col = r.Column.ToString();
                    r1c1 = "R" + r.Row + "C" + r.Column;
                    //MessageBox.Show("R1C1" + r1c1);
                    if (r.Value2 != null)
                    {
                        value2 = r.Value2.ToString();
                    }

                    if (r.Formula != null)
                    {
                        formula = r.Formula.ToString();
                    }

                    message = row + ":" + col + ":" + value2 + ":" + formula;
                    //MessageBox.Show("MESSAGE " + message);
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                return message;
            }

            public String getActiveCellRange()
            {
                object missing = Type.Missing;
                string cell = "";
                string col = "";
                string row = "";
                try
                {
                    Excel.Range r = Globals.ThisAddIn.Application.ActiveCell;

                    row = r.Row.ToString();
                    col = r.Column.ToString();
                    cell = r.Column + ":" + r.Row;
                    //MessageBox.Show("ID: " + r.ID + "value is :" + r.Text + " formula: " + r.Formula + "XPATH: " + r.XPath);
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    cell = "error: " + errorMsg;
                }

                //sets value using A1 notation - doesn't affect active cell
                Excel.Range r2 = Globals.ThisAddIn.Application.get_Range("A2", missing);
                r2.Value2 = "TEST";

                object ridx = 11;
                object cidx = 11;

                //sets value using r1c1 notation
                Excel.Range r3 = (Excel.Range)Globals.ThisAddIn.Application.Cells[ridx, cidx];
                r3.Value2 = "TEST2";

                //sets active cell using r1c1
                Excel.Range r4 = Globals.ThisAddIn.Application.ActiveCell;
                ridx = 2;
                cidx = 4;

                r4 = (Excel.Range)r4[ridx, cidx];

                r4.Activate();

                //MessageBox.Show(r4.get_Address(true, true, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, null, null));
                //MessageBox.Show("PAUSING");

                //sets active cell using a1
                Excel.Range r5 = Globals.ThisAddIn.Application.ActiveCell;
                r5 = r5.get_Range("B9", missing);
                r5.Activate();


                return cell;
            }

            //this actually replaces getActiveCell(), now use JSON
            public String getActiveCellText()
            {

                //move this to getRangeSelectionValues
                /*
             Excel.Range testr = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                foreach (Excel.Range cell in testr)
                {
                    MessageBox.Show("CELL VALUE IS: " + cell.Text);
                    string f = (string)cell.Formula;
                    string f2 = (string)cell.FormulaR1C1;
                        MessageBox.Show("Formula is: "+f);
                        MessageBox.Show("Formula 2 is :" + f2);

                        //cell.Formula = "=AVERAGE($A:1,$A:3)";
                        //cell.Calculate();
                }
                */


                string text = "";
                object missing = Type.Missing;
                try
                {
                    Excel.Range r = Globals.ThisAddIn.Application.ActiveCell;

                    text = "{ \"rowIdx\": " + "\"" + r.Row + "\""
                          + ",\"colIdx\": " + "\"" + r.Column + "\""
                          + ",\"coordinate\": " + "\"" + r.get_Address(r.Row, r.Column, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, missing, missing) + "\""
                          + ",\"value2\": " + "\"" + r.Value2 + "\""//r.value2
                          + ",\"formula\": " + "\"" + r.Formula + "\"" // r.Formula
                          + "}";

                    //text = r.Text + "";
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    text = "error: " + errorMsg;
                }

                // object missing = Type.Missing;

                //Here's how to do it, not sure of the use, stick with formula for cell at the moment
                // Excel.Worksheet w = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                // Excel.Range r = w.get_Range("A1", "B2");
                // r.Formula = "=AVERAGE(A1,B2)";
                // r.Calculate();

                return text;

            }
            //simple function, may be redundant as we have setCellValueA1
            public String setActiveCellValue(string value)
            {
                string message = "";
                try
                {
                    object txt = value;
                    Excel.Range r = Globals.ThisAddIn.Application.ActiveCell;
                    r.Value2 = txt;
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }
                return message;
            }
  
            public String setCellValueA1(string coordinate, string value, string sheetname)
            {
                //MessageBox.Show("setting for sheet: " + sheetname);
                object missing = Type.Missing;
                string message = "";

                try
                {
                    Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                
                    if (sheetname.Equals("active"))
                    {
                        Excel.Worksheet w = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                        Excel.Range r2 = w.get_Range(coordinate, missing);
                        r2.Value2 = value;
                    }
                    else
                    {
                        Excel.Worksheet w = (Excel.Worksheet)wb.Sheets[sheetname]; // (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[name];

                        Excel.Range r2 = w.get_Range(coordinate, missing);
                        r2.Value2 = value;
                    }
                  
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                    //MessageBox.Show("IN ERROR" + e.Message + "----" + e.StackTrace);
                }

                return message;
            }

            //utility, using so cell objects have both coordinate references
            public String convertA1ToR1C1(string coordinate)
            {
                string message = "";
                object missing = Type.Missing;
                try
                {
                    Excel.Range r2 = Globals.ThisAddIn.Application.get_Range(coordinate, missing);
                    message = r2.Column + ":" + r2.Row;
                }
                catch(Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }
                return message;
            }

            //utility, using so cell objects have both coordinate references
            public String convertR1C1ToA1(string rowIdx, string colIdx)
            {
                string message = "";
                object missing = Type.Missing;

                object r = Convert.ToInt32(rowIdx) - 1;
                object c = Convert.ToInt32(colIdx) - 1;

                try
                {
                    Excel.Range r2 = Globals.ThisAddIn.Application.get_Range("A1", missing);
                    r2 = r2.get_Offset(r, c);
                    message = r2.get_Address(r, c, Excel.XlReferenceStyle.xlA1, missing, missing);
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                return message;
            }

            public String clearWorksheet(string sheetName)
            {
                string message = "";
                object missing = Type.Missing;
                try
                {
                    Excel.Worksheet ws = null;

                    if (sheetName.Equals("active"))
                    {
                        ws = ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                    }
                    else
                    {
                        ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
                    }
                    
                    ws.Cells.Select();
                    ws.Cells.Clear();
                    Excel.Range r = (Excel.Range)ws.Cells[1, 1];
                    r.Select();
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                return message;

            }

            //separate function of try catches?
            public String getSheetType(string sheetName)
            {
                string sheetType = "";
                try
                {
                    Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

                    try
                    {

                        if ((Excel.Worksheet)wb.Worksheets[sheetName] is Excel.Worksheet)
                        {
                            sheetType = "xlWorksheet";
                            //     MessageBox.Show("HERE 1");
                        }
                    }
                    catch (Exception e)
                    {
                        string donothing_removewarning = e.Message;
                        try
                        {
                            if ((Excel.Chart)wb.Charts[sheetName] is Excel.Chart)  //wont work casting SH to worksheet chart in is
                            {
                                //       MessageBox.Show("HERE 3");
                                sheetType = "xlChart";
                                //      MessageBox.Show("HERE 2");

                            }
                        }
                        catch (Exception e2)
                        {
                            string donothing_removewarning2 = e2.Message;
                        }
                    }

                    //MessageBox.Show("what the hell is it?");
                    //not able to determine name
                    //some meaningful message here?
 
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    sheetType = "error: " + errorMsg;
                    MessageBox.Show("IN ERROR" + e.Message + "----" + e.StackTrace);
                }
                //check for other types, maybe update other function with types;
                //check in office 2007

                return sheetType;

            }

            public String getTempPath()
            {
                string tmpPath = "";
                try
                {
                    tmpPath = System.IO.Path.GetTempPath();
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    tmpPath = "error: " + errorMsg;
                }

                return tmpPath;
            }

            //UNDER HERE IS QUESTIONABLE
            static bool FileInUse(string path)
            {
                string __message = "";
                try
                {
                    //Just opening the file as open/create
                    using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                    {
                        //If required we can check for read/write by using fs.CanRead or fs.CanWrite
                    }
                    return false;
                }
                catch (IOException ex)
                {
                    //check if message is for a File IO
                    __message = ex.Message.ToString();
                    if (__message.Contains("The process cannot access the file"))
                        return true;
                    else
                        throw;
                }
            }

            public String saveActiveWorkbook(string path, string title, string url, string user, string pwd)
            {
                string message = "";
                object missing = Type.Missing;
                string newtitle = path + title;
                string tmptitle = path + "copyof_" + title;

                object t = newtitle;
                object tmpt = tmptitle;

                Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                try
                {
                    if (FileInUse(newtitle))
                    {
                        //in use
                        //need to save to copy, delete orig, save to orig, delete copy?
                        //lame, but may work til we come up with something else
                        if (wb.Name.Equals(title))
                        {
                            wb.SaveAs(tmpt, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
                            wb.Close(false, missing, missing);
                            File.Delete(newtitle);

                            Excel.Workbook wb2 = Globals.ThisAddIn.Application.Workbooks.Open(tmptitle, missing, false, missing, missing, missing, true, missing, missing, true, true, missing, missing, missing, missing);
                            wb2.SaveAs(t, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);

                            File.Delete(tmptitle);
                        }

                    }
                    else
                    {
                        wb.SaveAs(t, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
                    }
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                System.Net.WebClient Client = new System.Net.WebClient();
                Client.Headers.Add("enctype", "multipart/form-data");
                Client.Headers.Add("Content-Type", "application/octet-stream");

                try
                {
                    // FileStream fs = new FileStream(@"C:\Default.xlsx", FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                    FileStream fs = new FileStream(newtitle, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    int length = (int)fs.Length;
                    byte[] content = new byte[length];
                    fs.Read(content, 0, length);

                    try
                    {
                        Client.Credentials = new System.Net.NetworkCredential(user, pwd);
                        Client.UploadData(url, "POST", content);
                    }
                    catch (Exception e)
                    {
                        string errorMsg = e.Message;
                        message = "error: " + errorMsg;
                    }

                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                return message;
            }

            public String openXlsx(string path, string title, string url, string user, string pwd)
            {
               //MessageBox.Show("in the addin path:"+path+  "      title:"+title+ "   uri: "+url+" user: "+user+" pwd: "+pwd);
                string message = "";
                object missing = Type.Missing;
                string tmpdoc = "";
               
               // title=title.Replace("/","");

                try
                {
                    System.Net.WebClient Client = new System.Net.WebClient();
                    Client.Credentials = new System.Net.NetworkCredential(user, pwd);
                    tmpdoc = path + title;
                    //works thought path ends with / and doc starts with \ so you have C:tmp/\foo.xslx
                    //may need to fix
                    //MessageBox.Show("Tempdoc"+tmpdoc);
                    //Client.DownloadFile("http://w2k3-32-4:8000/test.xqy?uid=/Default.xlsx", tmpdoc);//@"C:\test2.xlsx");
                    Client.DownloadFile(url, tmpdoc);//@"C:\test2.xlsx");

                    //something weird with underscores, saves growth_model.xlsx becomes growth.model.xlsx
                    //may have to fix

                    Excel.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Open(tmpdoc, missing, false, missing, missing, missing, true, missing, missing, true, true, missing, missing, missing, missing);

                    /*
                     * another way 
                                        byte[] byteArray =  Client.DownloadData("http://w2k3-32-4:8000/test.xqy?uid=/Default.xlsx");//File.ReadAllBytes("Test.docx");
                                        using (MemoryStream mem = new MemoryStream())
                                        {

                                            mem.Write(byteArray, 0, (int)byteArray.Length);

                                            // using (OpenXmlPkg.SpreadsheetDocument sd = OpenXmlPkg.SpreadsheetDocument.Open(mem, true))
                                            // {
                                            // }
                        
                                            using (FileStream fileStream = new FileStream(@"C:\Test2.docx", System.IO.FileMode.CreateNew))
                                            {

                                                mem.WriteTo(fileStream);

                                            }

                       
                        
                                        }
                     * */

                    //OpenXmlPkg.SpreadsheetDocument xlPackage;
                    //xlPackage = OpenXmlPkg.SpreadsheetDocument.Open(strm, false);
                }
                catch (Exception e)
                {
                    //not always true, need to improve error handling or message or both
                    string origmsg = "A document with the name '"+title+"' is already open. You cannot open two documents with the same name, even if the documents are in different \nfolders. To open the second document, either close the document that's currently open, or rename one of the documents.";
                    MessageBox.Show(origmsg);
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                   
                }

                return message;
            }

            public String openXlsxWebDAV(string documenturi)
            {

                string message="";
                object missing = Type.Missing;
                object f = false;
                try
                {
                    //Excel.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    Excel.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Open(documenturi, missing, false, missing, missing, missing, true, missing, missing, true, true, missing, missing, missing, missing);
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }
                return message;

            }

            public String saveXlsxWebDAV(string title)
            {
                string message = "";
                object missing = Type.Missing;
                object t = title;
               Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
               try
               {
                   wb.SaveAs(t, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
               }
               catch (Exception e)
               {
                   string errorMsg = e.Message;
                   message = "error: " + errorMsg;
               }
  /*
    Object Filename,
    Object FileFormat,
    Object Password,
    Object WriteResPassword,
    Object ReadOnlyRecommended,
    Object CreateBackup,
    XlSaveAsAccessMode AccessMode,
    Object ConflictResolution,
    Object AddToMru,
    Object TextCodepage,
    Object TextVisualLayout,
    Object Local
  */
                    return message;
            }
            

            public String openDoc()
            {
                try
                {
                    object missing = Type.Missing;
                    object f = false;
                    Excel.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    wb = Globals.ThisAddIn.Application.Workbooks.Open("http://localhost:8011/openinml.xlsx", missing, false, missing, missing, missing, true, missing, missing, true, true, missing, missing, missing, missing);
                }
                catch (Exception e)
                {
                    MessageBox.Show("Error" + e.Message + "=====" + e.StackTrace);
                }

                return "foo";
            }

            //stubbed out, but not currently used. 
            public String setCellValueR1C1(int rowIndex, int colIndex, string value)
            {
                string message = "";
                return message;
            }

            //used for sna demo
            //but we may want some simple functions to insert csv into spreadsheet,
            //for those who don't want to create Cell objects, etc.
            public String insertRows(string edgelist1, string edgelist2, string vertices)
            {
                //Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                // Excel.Worksheet xls = null;
                Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

                //MessageBox.Show("ws index: " + ws.Index + "  ws name:" + ws.Name);
                //getX();

                //    ws = ( Excel.Worksheet)ws.Next;
                //    MessageBox.Show("ws index: " + ws.Index + "  ws name:" + ws.Name);
                //    ws = (Excel.Worksheet)ws.Previous;
                //    int start = 1;
                //    MessageBox.Show("ws index: " + ws.Index + "  ws name:" + ws.Name);

                //  string width = "B";
                // int length = 8; //determine by length of list

                string ppl1 = edgelist1; // "fred:fred:julie:tim:tim:frank:beth";
                string ppl2 = edgelist2; //"tim:julie:beth:beth:julie:fred:susan";
                string ppl3 = vertices;

                char[] delimiter = { ':' };
                string[] tmp1 = ppl1.Split(delimiter);
                string[] tmp2 = ppl2.Split(delimiter);
                string[] tmp3 = ppl3.Split(delimiter);

                int length1 = tmp1.Length;
                int length3 = tmp3.Length;

                //    for (int i = 0; i < ppl1.Length; i++)
                //   {
                //        MessageBox.Show("" + tmp1[i]);
                //   }

                int arrayind = 0;

                //  string startcol = "A1";
                // /string endcol = "A1";
                // Excel.Range rng = ws.get_Range(startcol, endcol);
                //  rng.Value2 = "PETE";

                //populate edges
                for (int i = 2; i < length1 + 2; i++)
                {
                    string startcol = "A" + i;
                    // string endcol = width + i;
                    string endcol = "A" + i;
                    Excel.Range rng = ws.get_Range(startcol, endcol);
                    foreach (Excel.Range cell in rng)
                    {
                        object x = tmp1[arrayind];

                        cell.Value2 = x;

                    }
                    arrayind++;
                }

                arrayind = 0;
                for (int i = 2; i < length1 + 2; i++)
                {
                    string startcol = "B" + i;
                    // string endcol = width + i;
                    string endcol = "B" + i;
                    Excel.Range rng = ws.get_Range(startcol, endcol);
                    foreach (Excel.Range cell in rng)
                    {
                        object x = tmp2[arrayind];

                        cell.Value2 = x;

                    }
                    arrayind++;
                }


                //populate vertices

                ws = (Excel.Worksheet)ws.Next;
                arrayind = 0;
                for (int i = 2; i < length3 + 2; i++)
                {
                    string startcol = "A" + i;
                    // string endcol = width + i;
                    string endcol = "A" + i;
                    Excel.Range rng = ws.get_Range(startcol, endcol);
                    foreach (Excel.Range cell in rng)
                    {
                        object x = tmp3[arrayind];

                        cell.Value2 = x;

                    }
                    arrayind++;
                }

                arrayind = 0;
                for (int i = 2; i < length3 + 2; i++)
                {
                    string startcol = "G" + i;
                    // string endcol = width + i;
                    string endcol = "G" + i;
                    Excel.Range rng = ws.get_Range(startcol, endcol);
                    foreach (Excel.Range cell in rng)
                    {
                        object x = tmp3[arrayind];

                        cell.Value2 = x;

                    }
                    arrayind++;
                }

                ws = (Excel.Worksheet)ws.Next;

                return "";
            }

            //not used, testing
            public String addCustomProperty(string key, string value)
            {
                string message = "";
                Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                object date = "2009-03-20";
                ws.CustomProperties.Add("date", date);

                return message;
            }

            public int getMacroCount()
            {
                int count=0;
                 try
                 {
                    VBIDE.VBProject proj = Globals.ThisAddIn.Application.ActiveWorkbook.VBProject;
                    count = proj.VBComponents.Count;
                 }
                 catch(Exception e)
                 {
                     MessageBox.Show(e.Message);
                 };

               return count;
            }

            public string getMacroName(int idx)
            {
                string message = "";
                try
                {
                    VBIDE.VBProject proj = Globals.ThisAddIn.Application.ActiveWorkbook.VBProject;
                    object o_idx = idx;
                    VBIDE.VBComponent vbComponent = proj.VBComponents.Item(o_idx);
                    message = vbComponent.Name;
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }
                return message;
            }

            public string getMacroProcedureName(int idx)
            {
                string message = "";
                try
                {
                    VBIDE.VBProject proj = Globals.ThisAddIn.Application.ActiveWorkbook.VBProject;
                    object o_idx = idx;
                    VBIDE.VBComponent vbComponent = proj.VBComponents.Item(o_idx);
                    if (vbComponent != null)
                    {
                        VBIDE.CodeModule componentCode = vbComponent.CodeModule;
                        if (componentCode.CountOfLines > 0)
                        {
                            int line = 1;
                            int componentCodeLines = componentCode.CountOfLines;
                            VBIDE.vbext_ProcKind procedureType = VBIDE.vbext_ProcKind.vbext_pk_Proc;
                            while (line < componentCodeLines)
                            {
                                string procedureName = componentCode.get_ProcOfLine(line, out procedureType);
                                //MessageBox.Show("procedure name" + procedureName);
                                if (procedureName != string.Empty)
                                {
                                    message = procedureName;
                                }
                                line++;
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }
                return message;
            }

            public string getMacroSignature(int idx)
            {
                string message = "";
                string signature = "";
                try
                {
                    VBIDE.VBProject proj = Globals.ThisAddIn.Application.ActiveWorkbook.VBProject;
                    object o_idx = idx;
                    VBIDE.VBComponent vbComponent = proj.VBComponents.Item(o_idx);
                    if (vbComponent != null)
                    {
                        VBIDE.CodeModule componentCode = vbComponent.CodeModule;
                        if (componentCode.CountOfLines > 0)
                        {
                          int line = 1;
                          int componentCodeLines = componentCode.CountOfLines;
                          VBIDE.vbext_ProcKind procedureType = VBIDE.vbext_ProcKind.vbext_pk_Proc;
                          while (line < componentCodeLines)
                          {
                            string procedureName = componentCode.get_ProcOfLine(line, out procedureType);
                            //MessageBox.Show("procedure name" + procedureName);
                            if (procedureName != string.Empty)
                            {
                                int procedureLines = componentCode.get_ProcCountLines(procedureName, procedureType);
                                int procedureStartLine = componentCode.get_ProcStartLine(procedureName, procedureType);
                                int codeStartLine = componentCode.get_ProcBodyLine(procedureName, procedureType);
                                int signatureLines = 1;
                                while (componentCode.get_Lines(codeStartLine, signatureLines).EndsWith("_"))
                                {
                                    signatureLines++;
                                }

                                signature = componentCode.get_Lines(codeStartLine, signatureLines);
                                signature = signature.Replace("\n", string.Empty);
                                signature = signature.Replace("\r", string.Empty);
                                signature = signature.Replace("_", string.Empty);
                                line += procedureLines - 1;
                            }
                            line++;
                          }
                        }
                        message = signature;
                    }
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                    MessageBox.Show("SIGNATURE ERROR: " + e.Message);
                }
                return message;
            }

            public string getMacroComments(int idx)
            {
                string message = "";
                string comments = "";
                try
                {
                    VBIDE.VBProject proj = Globals.ThisAddIn.Application.ActiveWorkbook.VBProject;
                    object o_idx = idx;
                    VBIDE.VBComponent vbComponent = proj.VBComponents.Item(o_idx);
                    if (vbComponent != null)
                    {
                        VBIDE.CodeModule componentCode = vbComponent.CodeModule;
                        if (componentCode.CountOfLines > 0)
                        {
                            int line = 1;
                            int componentCodeLines = componentCode.CountOfLines;
                            VBIDE.vbext_ProcKind procedureType = VBIDE.vbext_ProcKind.vbext_pk_Proc;
                            while (line < componentCodeLines)
                            {
                                string procedureName = componentCode.get_ProcOfLine(line, out procedureType);
                                //MessageBox.Show("procedure name" + procedureName);
                                if (procedureName != string.Empty)
                                {
                                    int procedureLines = componentCode.get_ProcCountLines(procedureName, procedureType);
                                    int procedureStartLine = componentCode.get_ProcStartLine(procedureName, procedureType);
                                    int codeStartLine = componentCode.get_ProcBodyLine(procedureName, procedureType);
                                    comments = "[No comments]";
                                    if (codeStartLine != procedureStartLine)
                                    {
                                        comments = componentCode.get_Lines(line, codeStartLine - procedureStartLine);
                                    }

                                }
                                line++;
                            }
                        }
                        message = comments;
                    }
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                    MessageBox.Show("COMMENTS ERROR: " + e.Message);
                }
                return message;
            }

            public string getMacroType(int idx)
            {
                 string message = "";
                 object o_idx = idx;
                 try
                 {
                     VBIDE.VBProject proj = Globals.ThisAddIn.Application.ActiveWorkbook.VBProject;
                     VBIDE.VBComponent vbComponent = proj.VBComponents.Item(o_idx);
                     message = vbComponent.Type+"";
                 }
                 catch (Exception e)
                 {
                     string errorMsg = e.Message;
                     message = "error: " + errorMsg;
                     MessageBox.Show("TYPE ERROR: " + e.Message);
                 }
                return message;
            }

            public string getMacroText(int idx)
            {
                VBIDE.VBProject proj = Globals.ThisAddIn.Application.ActiveWorkbook.VBProject;
               
                string componentFile = "";
                object o_idx = idx;
                try
                {
                    VBIDE.VBComponent vbComponent = proj.VBComponents.Item(o_idx);
                    
                    if (vbComponent != null)
                    {
                        VBIDE.CodeModule componentCode = vbComponent.CodeModule;

                        componentFile = "";
                        if (componentCode.CountOfLines > 0)
                        {
                            
                            for (int i = 0; i < componentCode.CountOfLines; i++)
                            {
                                componentFile += componentCode.get_Lines(i + 1, 1) + Environment.NewLine;
                            }
                        }
                       
                    }
                    
                }
                catch (Exception e)
                {
                    componentFile = "error: "+e.Message;
                    MessageBox.Show("ERROR" + componentFile);
                }
                    return componentFile;
            }

        public VBIDE.vbext_ComponentType getComponentTypeFromString(string componentType)
        {
            VBIDE.vbext_ComponentType type;

            if (componentType.Equals("vbext_ct_StdModule"))
            {
              type =   VBIDE.vbext_ComponentType.vbext_ct_StdModule;
            }
            else if (componentType.Equals("vbext_ct_ActiveXDesigner"))
            {
               type= VBIDE.vbext_ComponentType.vbext_ct_ActiveXDesigner;
            }
            else if (componentType.Equals("vbext_ct_Document"))
            {
                type = VBIDE.vbext_ComponentType.vbext_ct_Document;
            }
            else if (componentType.Equals("vbext_ct_Document"))
            {
              type=   VBIDE.vbext_ComponentType.vbext_ct_Document;
            }
            else if (componentType.Equals("vbext_ct_MSForm"))
            {
               type= VBIDE.vbext_ComponentType.vbext_ct_MSForm;
            }
            else
            {
               type = VBIDE.vbext_ComponentType.vbext_ct_StdModule;
            }
            
                return type;
        }

            public string removeMacro(string macroname)
            {
                string message = "";
                object id = macroname;
                int count = 0;
                try
                {
                    Globals.ThisAddIn.Application.EnableEvents = false;
                    VBIDE.VBProject proj = Globals.ThisAddIn.Application.ActiveWorkbook.VBProject;
                    //count = proj.VBComponents.Count;
                    foreach (VBIDE.VBComponent vbModule in proj.VBComponents)
                    {
                        VBIDE.CodeModule cm = vbModule.CodeModule;
                        if (cm.CountOfLines > 0)
                        {
                            int line = 1;
                            int componentCodeLines = cm.CountOfLines;
                            VBIDE.vbext_ProcKind procedureType = VBIDE.vbext_ProcKind.vbext_pk_Proc;
           
                            while (line < componentCodeLines)
                            {
                                string procedureName = cm.get_ProcOfLine(line, out procedureType);
                                if (procedureName.Equals(id))
                                {


                                    if (procedureName != string.Empty)
                                    {
                                        int procedureLines = cm.get_ProcCountLines(procedureName, procedureType);
                                        int procedureStartLine = cm.get_ProcStartLine(procedureName, procedureType);
                                        int codeStartLine = cm.get_ProcBodyLine(procedureName, procedureType);
                                        //MessageBox.Show("name" + procedureName + " lines" + procedureLines + " start" + procedureStartLine + " codeStart" + codeStartLine);
                                        vbModule.CodeModule.DeleteLines(procedureStartLine, procedureLines);
                                        break;
                                    }
                                   
                                    line++;
                                }
                            }
                        }
                    }
                    // proj.VBComponents.Remove(vbModule);
                    Globals.ThisAddIn.Application.EnableEvents = true;
                }
                catch (Exception e)
                {
                    Globals.ThisAddIn.Application.EnableEvents = true;
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                    MessageBox.Show("ERROR" + e.Message);
                }
                return message;
            }

            public string addMacro(string macro, string proctype) //addType
            {
                string message = "";
                try
                {
                    VBIDE.vbext_ComponentType type = getComponentTypeFromString(proctype);
                    
                    VBIDE.VBProject proj = Globals.ThisAddIn.Application.ActiveWorkbook.VBProject;
                    VBIDE.VBComponent vbModule = proj.VBComponents.Add(type);
                    string macroCode = macro; // "sub autoMacro()\r\n msgbox \" I am a macro. \" \r\n end sub";
                    vbModule.CodeModule.AddFromString(macroCode);
                   
                     
                    //oDoc.Application.Run("autoMacro", ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                }
                catch (Exception e)
                {
                    
                    //MessageBox.Show("ERROR " + e.Message);
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }
                //vbModule = null; 
                return message;
            }

            //ran in context of active sheet/missing are optional args to macro
            public string runMacro(string name) //ran in context of active sheet/missing are optional args to macro
            {
                string message = "";
                object missing = Type.Missing;

                try
                {
                    Globals.ThisAddIn.Application.Run(name, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);

                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }
                return message;
            }

            public string deletePicture(string sheetName, string imageName)
            {
                string message = "";
                object picName = imageName;

                try
                {
                    Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets[sheetName];
                    Excel.Picture pic = (Excel.Picture)ws.Pictures(picName);
                    pic.Delete();
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                    //MessageBox.Show(message);
                }
                return message;

            }

            public string exportChartImagePNG(string chartExportPath)
            {
                string message = "";
                try
                {
                    Excel.Chart c = Globals.ThisAddIn.Application.ActiveChart;

                    //name returns sheetname ' '(space) chartname
                    //by default, they don't have spaces, but you can give them spaces, so how to tease out chartname?
                    //chart.title? chart.index?

                    //Excel.Chart c = (Excel.Chart)Globals.ThisAddIn.Application.ActiveWorkbook.Charts[chartName];
                    //need to preserve clipboard before overwriting with image
                    //MessageBox.Show("2");
                    
                    //MessageBox.Show(chartExportPath);
                    c.Export(chartExportPath, "PNG", false);
                    
                    //c.CopyPicture(Microsoft.Office.Interop.Excel.XlPictureAppearance.xlScreen, Microsoft.Office.Interop.Excel.XlCopyPictureFormat.xlBitmap, Microsoft.Office.Interop.Excel.XlPictureAppearance.xlScreen);
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                    MessageBox.Show("ERROR CASTING CHART TO STRING: " + e.Message);
                }

           
                return message;
            }

            public string insertBase64ToImage(string base64String, string sheetName)
            {
       
                string message = "";
                try
                {
                    // Convert Base64 String to byte[]
                    byte[] imageBytes = Convert.FromBase64String(base64String);
                    MemoryStream ms = new MemoryStream(imageBytes, 0,
                      imageBytes.Length);

                    // Convert byte[] to Image
                    ms.Write(imageBytes, 0, imageBytes.Length);
                    Image image = Image.FromStream(ms, true);

                    Excel.Worksheet ws = null;

                    if (sheetName.Equals("active"))
                    {
                        ws = ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                    }
                    else
                    {
                        ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
                    }

                    Excel.Range oRange = (Excel.Range)ws.Cells[10, 10];

                    //backup clipboard
                    IDataObject bak = Clipboard.GetDataObject();
                    string text = "";
                    if (bak.GetDataPresent(DataFormats.Text))
                    {
                        text = (String)bak.GetData(DataFormats.Text);
                    }

                    System.Windows.Forms.Clipboard.SetDataObject(image, true);
                    ws.Paste(oRange, false);

                    Excel.Picture s = (Excel.Picture)Globals.ThisAddIn.Application.Selection;
                    message = s.Name;

                    if (!(text.Equals("")))
                        Clipboard.SetText(text);
               
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                    //MessageBox.Show("ERROR: " + e.Message);
                }


                return message;
            }
        
            public string base64EncodeImage(string chartPath)
            {
                string base64String = "";
           
                try
                {
                    Image img = Image.FromFile(chartPath);


                    //MemoryStream ms = new MemoryStream();
                    using (MemoryStream ms = new MemoryStream())
                    {
                        img.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        byte[] imageBytes = ms.ToArray();

                        // Convert byte[] to Base64 String
                        base64String = Convert.ToBase64String(imageBytes);

                    }

                    img.Dispose();

                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    base64String = "error: " + errorMsg;
                }
                
                return base64String;
            }

            public string deleteFile(string filePath)
            {
                string message = "";
                try
                {
                    File.Delete(filePath);
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }
                return message;
            }
    
    }
}
