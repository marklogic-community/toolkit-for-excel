using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.IO;

//using DocumentFormat.OpenXml.Packaging; //OpenXML sdk
using Office = Microsoft.Office.Core;
using Microsoft.Win32;
using Tools = Microsoft.Office.Tools.Excel;
using OpenXmlPkg = DocumentFormat.OpenXml.Packaging;


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
            private string addinVersion = "@MAJOR_VERSION.@MINOR_VERSION@PATCH_VERSION";
            HtmlDocument htmlDoc;

            public UserControl1()
            {
                InitializeComponent();
                //bool regEntryExists = checkUrlInRegistry();
                webUrl = ac.getWebURL();
                //MessageBox.Show("URL" + webUrl);

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
                //set for sureo once user sets color scheme manually from button
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
                    Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                    wsname = ws.Name;
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    wsname = "error: " + errorMsg;
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
                    //ws.Activate();
                    //see note above (function sig), name defaults to Sheet1
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    wbname = "error: " + errorMsg;
                }
                return wbname;
            }

            /* public String addWorksheet()
             {
                 //add default, return name of added ws
                 string ws = "FOO";
                 return ws;
             }
             * */

            //have to account for empty string "", what's up with overloaded functions not working?
            public String addWorksheet(string name)//#sheets as param?
            {
                string message = "";
                Excel.Worksheet ws = null;
                try
                {
                    object missing = Type.Missing;
                    int count = 1;  //see note in function sig
                    // object after = "Sheet2";
                    // Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add(missing,missing,missing,missing);
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
            public String setCell(string coordinate, string value, string name)
            {
                object missing = Type.Missing;
                string message = "";

                try
                {
                    Excel.Workbook wb = Globals.ThisAddIn.Application.Workbooks["Book1"];
                    Excel.Worksheet w = (Excel.Worksheet)wb.Sheets[name]; // (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[name];
                    Excel.Range r2 = w.get_Range(coordinate, missing);
                    r2.Value2 = value;
                }
                catch (Exception e)
                {
                    MessageBox.Show("IN ERROR" + e.Message + "----" + e.StackTrace);
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

                    //Globals.ThisAddIn.Application.Workbooks[name].Activate();
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                return message;
            }

            public String getTempPath()
            {
                string tmpPath = System.IO.Path.GetTempPath();
               //MessageBox.Show("returning "+tmpPath);
                return tmpPath;
            }
            public String openXlsxWebDAV(string documenturi)
            {
                //MessageBox.Show("IN ADDIN");
                string message="";
                object missing = Type.Missing;
                object f = false;
                //Excel.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Open(documenturi, missing, false, missing, missing, missing, true, missing, missing, true, true, missing, missing, missing, missing);
                   
                return message;

            }

            public String openXlsx2(string path, string title, string url, string user, string pwd)
            {
                // MessageBox.Show("in the addin filename:"+filename+ "   uri: "+uri);
                string message = "";
                object missing = Type.Missing;
                Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

                try
                {
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                      System.Net.WebClient Client = new System.Net.WebClient();
                      //Client.Credentials = new System.Net.NetworkCredential("zeke", "zeke");
                      Client.Credentials = new System.Net.NetworkCredential(user, pwd);
                      //string tmppath = getTempPath();
                      string tmpdoc = path + title;
                      //Client.DownloadFile("http://w2k3-32-4:8000/test.xqy?uid=/Default.xlsx", tmpdoc);//@"C:\test2.xlsx");
                      byte[] byteArray = Client.DownloadData(url);
                      //Excel.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Open(tmpdoc, missing, false, missing, missing, missing, true, missing, missing, true, true, missing, missing, missing, missing);
                      memoryStream.Write(byteArray, 0, byteArray.Length);

                        //no way to open Excel from memory stream at this point, can do this in Word/powerpoint by serializing as base64 string, or just returning xml for document
                        //and using insertXML
                      using (OpenXmlPkg.SpreadsheetDocument xldoc = OpenXmlPkg.SpreadsheetDocument.Open(memoryStream, true))
                      {
                        
                      }
                      

                       
                    }
                     
                    //OpenXmlPkg.SpreadsheetDocument sd
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
                    MessageBox.Show("problem" + e.Message + "   " + e.StackTrace);
                }


                return message;
            }


            public String openXlsx(string path, string title, string url, string user, string pwd)
            {
                // MessageBox.Show("in the addin filename:"+filename+ "   uri: "+uri);
                string message = "";
                object missing = Type.Missing;

                try
                {
                    System.Net.WebClient Client = new System.Net.WebClient();
                    //Client.Credentials = new System.Net.NetworkCredential("zeke", "zeke");
                    Client.Credentials = new System.Net.NetworkCredential(user, pwd);
                    //string tmppath = getTempPath();
                    string tmpdoc = path + title;
                    //Client.DownloadFile("http://w2k3-32-4:8000/test.xqy?uid=/Default.xlsx", tmpdoc);//@"C:\test2.xlsx");
                    Client.DownloadFile(url, tmpdoc);//@"C:\test2.xlsx");
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
                    MessageBox.Show("problem" + e.Message + "   " + e.StackTrace);
                }


                return message;
            }

            public String openXlsx(string filename, string uri)
            {
               // MessageBox.Show("in the addin filename:"+filename+ "   uri: "+uri);
                string message = "";
                object missing = Type.Missing;

                try
                {
                    System.Net.WebClient Client = new System.Net.WebClient();
                    //Client.Credentials = new System.Net.NetworkCredential("oslo", "oslo");
                    Client.Credentials = new System.Net.NetworkCredential("zeke", "zeke");
                    string tmppath = getTempPath();
                    string tmpdoc = tmppath + filename;
                    //Client.DownloadFile("http://w2k3-32-4:8000/test.xqy?uid=/Default.xlsx", tmpdoc);//@"C:\test2.xlsx");
                    Client.DownloadFile("http://localhost:8000/test.xqy?uid="+uri, tmpdoc);//@"C:\test2.xlsx");
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
                    MessageBox.Show("problem"+e.Message+"   "+e.StackTrace);
                }


                return message;
            }
/*
            public String saveActiveWorkbook(string path, string title, string url, string user, string pwd)
            {
                object missing = Type.Missing;
                string newtitle = path + title;
                MessageBox.Show("NEW PATH" + newtitle);
                object t = title;
                Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                try
                {
                    wb.SaveAs(t, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);

                }
                catch (Exception e)
                {
                    MessageBox.Show("that shizz didn't work?!?!?" + e.Message + "===" + e.StackTrace);
                }

                System.Net.WebClient Client = new System.Net.WebClient();
                //Client.Headers.AllKeys;
                Client.Headers.Add("enctype", "multipart/form-data");
                //Client.Headers.Add("Content-Type","application/x-www-form-urlencoded");
                Client.Headers.Add("Content-Type", "application/octet-stream");

                //Client.Headers.Add("Content-Transfer-Encoding","application/octet-stream");

                try
                {

                    // FileStream fs = new FileStream(@"C:\Default.xlsx", FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                    FileStream fs = new FileStream(title, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    int length = (int)fs.Length;
                    byte[] content = new byte[length];
                    fs.Read(content, 0, length);




                    try
                    {
                        // MessageBox.Show("URL: " + url);
                        // MessageBox.Show("TITLE " + title);
                        //Client.Credentials = new System.Net.NetworkCredential("oslo", "oslo");
                        Client.Credentials = new System.Net.NetworkCredential(user, pwd);
                        //Client.UploadFile("http://localhost:8000/addinSampleExcelMeta/upload2.xqy?uid=Default.xlsx", "POST", @"c:\Default.xlsx");//@"c:\tmp.xml");

                        //Client.UploadData("http://localhost:8000/addinSampleExcelMeta/upload2.xqy?uid=Default.xlsx", "POST",content);
                        Client.UploadData(url, "POST", content);
                        //Client.UploadFile(url, "POST", title);
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("ERROR" + e.Message + "      " + e.StackTrace);
                    }

                }
                catch (Exception e)
                {
                    MessageBox.Show("THIS BLEW UP:" + e.Message + "                 " + e.StackTrace);
                }
                return "foo";
            }
*/

            public String saveActiveWorkbook(string path, string title, string url, string user, string pwd)
            {
                object missing = Type.Missing;
                string newtitle = path + title;
                MessageBox.Show("NEW PATH" + newtitle);
                string tmptitle = "copyof_" + title;

                object t = newtitle;
                object tmpt = tmptitle;

                Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                try
                {
                    wb.SaveAs(t, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);

                }
                catch (Exception e)
                {
                    MessageBox.Show("that shizz didn't work?!?!?" + e.Message + "===" + e.StackTrace);
                }

                System.Net.WebClient Client = new System.Net.WebClient();
                //Client.Headers.AllKeys;
                Client.Headers.Add("enctype", "multipart/form-data");
                //Client.Headers.Add("Content-Type","application/x-www-form-urlencoded");
                Client.Headers.Add("Content-Type", "application/octet-stream");

                //Client.Headers.Add("Content-Transfer-Encoding","application/octet-stream");

                try
                {

                    // FileStream fs = new FileStream(@"C:\Default.xlsx", FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                    FileStream fs = new FileStream(newtitle, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    int length = (int)fs.Length;
                    byte[] content = new byte[length];
                    fs.Read(content, 0, length);




                    try
                    {
                        // MessageBox.Show("URL: " + url);
                        // MessageBox.Show("TITLE " + title);
                        //Client.Credentials = new System.Net.NetworkCredential("oslo", "oslo");
                        Client.Credentials = new System.Net.NetworkCredential(user, pwd);
                        //Client.UploadFile("http://localhost:8000/addinSampleExcelMeta/upload2.xqy?uid=Default.xlsx", "POST", @"c:\Default.xlsx");//@"c:\tmp.xml");

                        //Client.UploadData("http://localhost:8000/addinSampleExcelMeta/upload2.xqy?uid=Default.xlsx", "POST",content);
                        Client.UploadData(url, "POST", content);
                        //Client.UploadFile(url, "POST", title);
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("ERROR" + e.Message + "      " + e.StackTrace);
                    }

                }
                catch (Exception e)
                {
                    MessageBox.Show("THIS BLEW UP:" + e.Message + "                 " + e.StackTrace);
                }
                return "foo";
            }

            public String saveXlsx(string title, string url)
            {
                object missing = Type.Missing;
                object t = title;
                Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                try
                {
                    wb.SaveAs(t, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
                    
                }
                catch (Exception e)
                {
                    MessageBox.Show("that shizz didn't work?!?!?" + e.Message + "===" + e.StackTrace);
                }

                System.Net.WebClient Client = new System.Net.WebClient();
                //Client.Headers.AllKeys;
                Client.Headers.Add("enctype", "multipart/form-data");
               //Client.Headers.Add("Content-Type","application/x-www-form-urlencoded");
                Client.Headers.Add("Content-Type", "application/octet-stream");

              //Client.Headers.Add("Content-Transfer-Encoding","application/octet-stream");

               try
               {
                   
                  // FileStream fs = new FileStream(@"C:\Default.xlsx", FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                   FileStream fs = new FileStream(title, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                   int length = (int)fs.Length;
                   byte[] content = new byte[length];
                   fs.Read(content, 0, length);




                   try
                   {
                      // MessageBox.Show("URL: " + url);
                      // MessageBox.Show("TITLE " + title);
                       //Client.Credentials = new System.Net.NetworkCredential("oslo", "oslo");
                       Client.Credentials = new System.Net.NetworkCredential("zeke", "zeke");
                       //Client.UploadFile("http://localhost:8000/addinSampleExcelMeta/upload2.xqy?uid=Default.xlsx", "POST", @"c:\Default.xlsx");//@"c:\tmp.xml");
                       
                       //Client.UploadData("http://localhost:8000/addinSampleExcelMeta/upload2.xqy?uid=Default.xlsx", "POST",content);
                       Client.UploadData(url, "POST", content);
                       //Client.UploadFile(url, "POST", title);
                   }
                   catch (Exception e)
                   {
                       MessageBox.Show("ERROR" + e.Message + "      " + e.StackTrace);
                   }

               }
               catch (Exception e)
               {
                   MessageBox.Show("THIS BLEW UP:" +e.Message+"                 "+ e.StackTrace);
               }
                return "foo";
            }

            public String saveXlsxWebDAV(string title)
            {
                string message = "";
                object missing = Type.Missing;
                //string  tmp = System.IO.Path.GetTempPath(); 
                MessageBox.Show("document: "+title);
                object t = title;
               Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
               try
               {
                   wb.SaveAs(t, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing);
               }
               catch (Exception e)
               {
                   MessageBox.Show("that shizz didn't work?!?!?" + e.Message + "===" + e.StackTrace);
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
                               // Excel.Application excelApp;
                               // excelApp = new Microsoft.Office.Interop.Excel.Application();
                               // excelApp.Visible = true;
                    /*these 2 work, but overwrite existing try add new workbook first below
                    Excel.Workbook wb =  Globals.ThisAddIn.Application.Workbooks.Open("http://localhost:8011/openinml.xlsx", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                    wb.Activate();
                    */
                               // excelApp.Workbooks.Open("http://localhost:8011/openinml.xlsx", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);



                    //Excel.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    object f = false;
                    Excel.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    wb = Globals.ThisAddIn.Application.Workbooks.Open("http://localhost:8011/openinml.xlsx", missing, false, missing, missing, missing, true, missing, missing, true, true, missing, missing, missing, missing);
                   
                    //Excel.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                   //wb.ChangeFileAccess(Excel.XlFileAccess.xlReadWrite, missing, true);
                    
                   // Excel.Workbook wb2 = Globals.ThisAddIn.Application.ActiveWorkbook;
                   // wb2.ChangeFileAccess(Excel.XlFileAccess.xlReadWrite, missing, true);
                  
                    //Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                    //app.Workbooks.Open("http://localhost:8011/openinml.xlsx", missing, false, missing, missing, missing, true, missing, missing, true, true, missing, missing, missing, missing);
                    //wb = app.Workbooks[1];
                   // wb.Activate();

                   


                    



                }
                catch (Exception e)
                {
                    MessageBox.Show("Error" + e.Message + "=====" + e.StackTrace);
                }

                return "foo";
            }

            //TESTER
            //functions we may want
            //clearNamedRange
            //addComment
            //clear workbook
            //clear sheet
            //clear range

            public String clearActiveWorksheet()
            {
                string message = "FOO";
                object missing = Type.Missing;

                //could do it by name, but then do we reset if user on different sheet?
                //plus, we have other functions for getting/setting active worksheet
                //can just loop through to delete all contents

                //Excel.Worksheet ws =   (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[name]; // ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[name]).Select( missing);
                Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                //ws.Select(missing);
                ws.Cells.Select();
                ws.Cells.Clear();
                Excel.Range r = (Excel.Range)ws.Cells[1, 1];
                r.Select();

                return message;

            }


            public String addAutoFilter(string coordinate1, string coordinate2, string criteria1, string v_operator, string criteria2)
            {
                MessageBox.Show("c1: " + coordinate1 + " c2: " + coordinate2 + " crit1: " + criteria1 + " op: " + v_operator + " crit2: " + criteria2);
                string message = "";
                object missing = Type.Missing;

                Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                Excel.Range rg = ws.get_Range(coordinate1, coordinate2);

                //rg.AutoFilter(1,missing,Excel.XlAutoFilterOperator.xlAnd,missing, true);
                rg.AutoFilter(1, "<>", Excel.XlAutoFilterOperator.xlOr, missing, true);

                return message;


            }

            public String addNamedRange(string coordinate1, string coordinate2, string rngName)
            {
                //MessageBox.Show(coordinate1 + " " + coordinate2);
                string message = "";
                object missing = Type.Missing;

                //add check for name
                //get range, check name

                Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                Excel.Range rg = ws.get_Range(coordinate1, coordinate2);
                Excel.Name nm = ws.Names.Add(rngName, rg, true, missing, missing, missing, missing, missing, missing, missing, missing);

                //test autofilter
                //Excel.XlAutoFilterOperator.xlAnd, missing, true

                //this works, have to break out into own function, also, clearCells for named range removes
                //"<>" selects all non empty
                //otherwise, use criteria
                // offset, criteria1, operator (and/or), criteria2,visibledropdown
                //rg.AutoFilter(1,missing,Excel.XlAutoFilterOperator.xlAnd,missing, true);
                // rg.AutoFilter(1, "22", Excel.XlAutoFilterOperator.xlAnd, "23", true);
                // rg.AutoFilter(1, "Sue", Excel.XlAutoFilterOperator.xlAnd, missing, true);



                // MessageBox.Show("NAME: "+nm.Name + "| Refers to: " + nm.RefersTo.ToString());

                /*  string names = "There are " + ws.Names.Count + " names: ";

                  Excel.Names ns = Globals.ThisAddIn.Application.ActiveWorkbook.Names;
                  MessageBox.Show(ns.Count+"");

                  foreach(Excel.Name n in ns)
                      MessageBox.Show("NAMES ARE: " + n.Name);

                  */


                /*System.Collections.IEnumerator en = ws.Names.GetEnumerator();
                while (en.MoveNext())
                {
                    object nameObj = en.Current;
                    Excel.Name name = nameObj as Excel.Name;
                    if (name != null)
                        names += name.Name;
                }*/
                // Excel.Names ns = ws.Names;
                // foreach(Excel.Name n in ns)


                return message;
            }

            public String getNamedRangeRangeNames()
            {
                string message = "";
                string names = "";

                // Excel.Names nnn = Globals.ThisAddIn.Application.Names;
                // foreach (Excel.Name x in nnn)
                //   MessageBox.Show("name is " + x.Name);

                Excel.Names ns = Globals.ThisAddIn.Application.ActiveWorkbook.Names;
                //Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                //Excel.Names ns = ws.Names;

                foreach (Excel.Name n in ns)
                    names += n.Name + ":";

                names = names.Substring(0, names.Length - 1);

                message = names;
                return message;

            }

            //remove named range?

            public String removeNamedRange(string rngName)
            {
                string message = "";
                object missing = Type.Missing;

                Excel.Names ns = Globals.ThisAddIn.Application.ActiveWorkbook.Names;
                foreach (Excel.Name nDel in ns)
                {
                    // MessageBox.Show("NAME" + nDel.Name);
                    if (nDel.Name.EndsWith(rngName))
                    {
                        //MessageBox.Show("deleting name");
                        nDel.Delete();
                    }
                }


                return message;
            }

            public String clearNamedRange(string rngName)
            {
                String message = "";
                // int ctlIndex = Globals.Sheet1.Controls.IndexOf(nm.NameLocal.ToString());
                //object ctl = Globals.Sheet1.Controls[ctlIndex];
                //Microsoft.Office.Tools.Excel.NamedRange nr = ctl as Microsoft.Office.Tools.Excel.NamedRange;
                //nr.Delete();
                /////\

                //Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                object missing = Type.Missing;

                //try using names first
                //Excel.Sheets ws = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets;//Globals.ThisAddIn.Application.Worksheets;
                String names = getActiveWorkbookWorkSheetNames();

                //how to set active worksheet sheet
                ///((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[name]).Select(missing);

                //int ctlIdx = Globals.ThisAddIn.DataBindings.Control.Controls.
                //Globals.ThisAddIn.Application.ThisWorkbook.
                //object ctl = Globals.ThisAddIn.DataBindings.Control.Controls[ctlIdx];
                //Tools.NamedRange nr = ctl as Tools.NamedRange;
                //nr.Delete(); 

                Excel.Range r = null;

                //loop thru all sheets til we find range, return first, else, give up
                //names have to be unique, so this seems like a safe bet
                // foreach (Excel.Worksheet n in ws)
                char x = '|';
                foreach (String name in names.Split(x))
                {
                    // MessageBox.Show("NAME " + name);

                    setActiveWorksheet(name);
                    Excel.Worksheet n = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

                    try
                    {
                        // MessageBox.Show("IN TRY");
                        r = n.get_Range(rngName, missing);
                        if (r != null)
                        {
                            //Excel.Name nn = (Excel.Name)r.Name;
                            //MessageBox.Show("NAME: "+nn.Name);
                            r.Select();
                            r.Clear();
                            //r.Name = "";

                            /*     Excel.Names ns = Globals.ThisAddIn.Application.ActiveWorkbook.Names;
                                 foreach (Excel.Name nDel in ns)
                                 {
                                     MessageBox.Show("NAME" + nDel.Name);
                                     if (nDel.Name.EndsWith(rngName))
                                     {
                                         MessageBox.Show("deleting name");
                                         nDel.Delete();
                                     }
                                 }
                             * */
                            //r.Name="";
                            //r.Name = "";
                            // Tools.NamedRange nr = (Tools.NamedRange)r.Name;
                            //nr.Delete();
                            //nr.Delete();

                            //Excel.Name nnn = (Excel.Name)r.Name;
                            //MessageBox.Show(" HERE"+nnn.Name);
                            //r.Name = "";



                            break;
                        }
                    }
                    catch (Exception e)
                    {
                        //MessageBox.Show("IN CATCH"+e.Message+e.StackTrace);
                        r = null;
                    }


                }

                return message;
            }

            public String setActiveRangeByName(string rngName)
            {
                String message = "";

                //first get activeworksheet name (don't need this, unless we reset to original page
                //Excel.Worksheet w = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                //MessageBox.Show("NAME OF SHEET"+w.Name+"NAME OF RANGE "+rngName);
                //get all worksheet names
                object missing = Type.Missing;
                Excel.Sheets ws = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets;//Globals.ThisAddIn.Application.Worksheets;
                Excel.Range r = null;

                //loop thru all sheets til we find range, return first, else, give up
                //names have to be unique, so this seems like a safe bet
                foreach (Excel.Worksheet n in ws)
                {
                    string wsname = n.Name;
                    setActiveWorksheet(wsname);
                    try
                    {
                        // MessageBox.Show("IN TRY");
                        r = n.get_Range(rngName, missing);
                        if (r != null)
                        {
                            r.Activate();
                            break;
                        }
                    }
                    catch
                    {
                        MessageBox.Show("IN CATCH");
                        r = null;
                    }


                }

                return message;
            }

            public String clearRange(string startcoord, string endcoord)
            {
                MessageBox.Show("HERE");
                string message = "";
                object missing = Type.Missing;
                Excel.Worksheet w = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;//Globals.ThisAddIn.Application.Worksheets;
                Excel.Range r = w.get_Range(startcoord, endcoord);
                r.Clear();

                return message;

            }

            public String getSelectedRangeCoordinates()
            {
                string message = "";
                string firstCellCoordinate = "";
                string lastCellCoordinate = "";

                Excel.Range r = (Excel.Range)Globals.ThisAddIn.Application.Selection;

                int start = 1;
                int end = r.Count;
                int count = 1;
                MessageBox.Show("COUNT" + r.Count);

                foreach (Excel.Range r2 in r)
                {
                    //r4.get_Address(true, true, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, null, null)
                    //MessageBox.Show("" + r2.Column + r2.Row);

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
                MessageBox.Show("RANGE: " + message);
                return message;
            }

            public String getSelectedCells()
            {
                MessageBox.Show("IN FUNCTION");
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

                        /*               try
                                          {
                                              if (rng.Name!=null)
                                              {
                                                  Excel.Name name = r.Name as Excel.Name;
                                                  MessageBox.Show(name.Name);
                                              }
                                          }
                                          catch (Exception exx)
                                          {
                                              MessageBox.Show("ERROR"+exx.Message+"      "+exx.StackTrace);
                                          }
                       */
                        // MessageBox.Show("NAME: "+r.Name + "");

                        string cell = "";
                        cell = "{ \"rowIdx\": " + "\"" + row + "\""
                             + ",\"colIdx\": " + "\"" + col + "\""
                             + ",\"coordinate\": " + "\"" + coordinate + "\""
                             + ",\"value2\": " + "\"" + value2 + "\""//r.value2
                             + ",\"formula\": " + "\"" + formula + "\"" // r.Formula
                             + "}";

                        cells += cell + ",";

                        // MessageBox.Show("CELL: "+cell);
                    }

                    cells = cells.Substring(0, cells.Length - 1);
                    cells += "]";
                    MessageBox.Show("message: " + cells);

                    message = cells;

                    /*   row = r.Row.ToString();
                       col = r.Column.ToString();
                       //cell = r.Column + ":" + r.Row;
                       r1c1 = "R" + r.Row + "C" + r.Column;
                       MessageBox.Show("R1C1" + r1c1);
                       //coordinate=convertR1C1ToA1(r.Row.ToString(),r.Column.ToString());
                       if (r.Value2 != null)
                       {
                           value2 = r.Value2.ToString();
                       }

                       if (r.Formula != null)
                       {
                           formula = r.Formula.ToString();
                       }

                       message = row + ":" + col + ":" + value2 + ":" + formula;
                       MessageBox.Show("MESSAGE " + message);
                       // MessageBox.Show("coordinate: "+coordinate+" r1c1: "+r1c1+" ID: " + r.ID + " value2: " + r.Value2 + " formula: " + r.Formula + " XPATH: " + r.XPath);
                   */
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
                string coordinate = "";
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
                    //cell = r.Column + ":" + r.Row;
                    r1c1 = "R" + r.Row + "C" + r.Column;
                    MessageBox.Show("R1C1" + r1c1);
                    //coordinate=convertR1C1ToA1(r.Row.ToString(),r.Column.ToString());
                    if (r.Value2 != null)
                    {
                        value2 = r.Value2.ToString();
                    }

                    if (r.Formula != null)
                    {
                        formula = r.Formula.ToString();
                    }

                    message = row + ":" + col + ":" + value2 + ":" + formula;
                    MessageBox.Show("MESSAGE " + message);
                    // MessageBox.Show("coordinate: "+coordinate+" r1c1: "+r1c1+" ID: " + r.ID + " value2: " + r.Value2 + " formula: " + r.Formula + " XPATH: " + r.XPath);
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
                    MessageBox.Show("ID: " + r.ID + "value is :" + r.Text + " formula: " + r.Formula + "XPATH: " + r.XPath);
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

                MessageBox.Show(r4.get_Address(true, true, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, null, null));

                MessageBox.Show("PAUSING");

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

            //utility, using so cell objects have both coordinate references
            public String convertA1ToR1C1(string coordinate)
            {
                string message = "";
                object missing = Type.Missing;
                Excel.Range r2 = Globals.ThisAddIn.Application.get_Range(coordinate, missing);
                //string test = r2.get_Address(missing, missing, Excel.XlReferenceStyle.xlR1C1, missing,missing);
                //MessageBox.Show("TEST A1toR1C1: " + test);
                message = r2.Column + ":" + r2.Row;
                return message;
            }

            //utility, using so cell objects have both coordinate references
            public String convertR1C1ToA1(string rowIdx, string colIdx)
            {

                //string r1c1="R"+rowIdx+"C"+colIdx;
                //MessageBox.Show("R1C1="+r1c1);
                //try get_Range with r1C1 (type it out R1C1: etc. 
                string message = "";
                object missing = Type.Missing;

                object r = Convert.ToInt32(rowIdx) - 1;
                object c = Convert.ToInt32(colIdx) - 1;

                try
                {
                    Excel.Range r2 = Globals.ThisAddIn.Application.get_Range("A1", missing);
                    r2 = r2.get_Offset(r, c);
                    message = r2.get_Address(r, c, Excel.XlReferenceStyle.xlA1, missing, missing);

                    //MessageBox.Show("TEST 2: " + test);
                }
                catch (Exception e)
                {
                    MessageBox.Show("ERROR :" + e.Message);
                }

                //MessageBox.Show("RETURNING" + message);
                return message;
            }

            //stubbed out, but not currently used. 
            public String setCellValueR1C1(int rowIndex, int colIndex, string value)
            {
                string message = "";
                return message;
            }

            //how we set cell values currently
            //may want to use entire cell object
            public String setCellValueA1(string coordinate, string value)
            {
                //MessageBox.Show("IN STRING METHOD");
                object missing = Type.Missing;
                string message = "";

                try
                {
                    Excel.Range r2 = Globals.ThisAddIn.Application.get_Range(coordinate, missing);
                    r2.Value2 = value;
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }
                return message;
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

            //used for sna demo
            //but we may want some simple functions to insert csv into spreadsheet,
            //for those who don't want to create Cell objects, etc.
            public String insertRows(string edgelist1, string edgelist2, string vertices)
            {
                //  MessageBox.Show("TESTING FROM WITHIN ADDIN");
                //Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                // Excel.Worksheet xls = null;
                Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

                MessageBox.Show("ws index: " + ws.Index + "  ws name:" + ws.Name);
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
                /*    for (int i = 2; i < length + 2; i++)
                    {
                        string startcol = "A" + i;
                        string endcol = width + i;
                        Excel.Range rng = ws.get_Range(startcol, endcol);
                        foreach (Excel.Range cell in rng)
                        {
                            cell.Value2 = 3.3;
                        }
                    }
                    */

                //   Excel.Range rng = ws.get_Range("A1", "B1");

                /*     foreach (Excel.Range cell in rng){

                         cell.Value2 = 3.3;
                     }

                     Excel.Range rng2 = ws.get_Range("B2", "C2");
                     foreach (Excel.Range cell in rng2)
                     {

                         cell.Value2 = 3.3;
                     }
                     */



                return "";
            }

        
    
    }
}
