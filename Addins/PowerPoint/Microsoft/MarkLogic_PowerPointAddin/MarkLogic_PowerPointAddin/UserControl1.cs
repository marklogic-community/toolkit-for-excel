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
 * UserControl1.cs - the api called from MarkLogicPowerPointAddin.js.  The methods here map directly to functions in the .js.
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
using PwrPt = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.IO;
//using DocumentFormat.OpenXml.Packaging; //OpenXML sdk
using Office = Microsoft.Office.Core;
using Microsoft.Win32;
using PPT = Microsoft.Office.Interop.PowerPoint;
using OX = DocumentFormat.OpenXml.Packaging;
using System.Web.Script.Serialization;



namespace MarkLogic_PowerPointAddin
{
    [ComVisible(true)]
    public partial class UserControl1 : UserControl
    {
        private AddinConfiguration ac = AddinConfiguration.GetInstance();
        private string webUrl = "";
        private bool debug = false;
        private string color = "";
        private string addinVersion = "@MAJOR_VERSION.@MINOR_VERSION@PATCH_VERSION";
        HtmlDocument htmlDoc;

        public UserControl1()
        {
            InitializeComponent();
           // bool regEntryExists = checkUrlInRegistry();
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

        public enum ColorScheme : int
        {
            Blue = 1,
            Silver = 2,
            Black = 3,
            Unknown = 4
        };

        public ColorScheme TryGetColorScheme()
        {

            ColorScheme CurrentColorScheme = 0;
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
                PwrPt.Presentation pres = Globals.ThisAddIn.Application.ActivePresentation;
                int count = pres.CustomXMLParts.Count;

                foreach (Office.CustomXMLPart c in pres.CustomXMLParts)
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
                PwrPt.Presentation pres = Globals.ThisAddIn.Application.ActivePresentation;
                Office.CustomXMLPart cx = pres.CustomXMLParts.SelectByID(id);

                if (cx != null)
                    custompiecexml = cx.XML;

                /*another way (used until I discovered SelectByID(id) above)
                  keeping here for notes, but this is for Word, translate to XL
                  foreach (Office.CustomXMLPart c in doc.CustomXMLParts)
                  {
                      if (c.BuiltIn.Equals(false) && c.Id.Equals(id))
                      {
                          Office.CustomXMLNode x = c.DocumentElement;
                          custompiecexml = x.XML;
                      }
                
                  }
                 */
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
                PwrPt.Presentation pres = Globals.ThisAddIn.Application.ActivePresentation;
                Office.CustomXMLPart cx = pres.CustomXMLParts.Add(String.Empty, new Office.CustomXMLSchemaCollectionClass());
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
                PwrPt.Presentation pres = Globals.ThisAddIn.Application.ActivePresentation;
                foreach (Office.CustomXMLPart c in pres.CustomXMLParts)
                {
                    if (c.BuiltIn.Equals(false) && c.Id.Equals(id))
                    {
                        //Office.CustomXMLNode x = c.DocumentElement;
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

        public Image byteArrayToImage(byte[] byteArrayIn)
        {
            try
            {
                MemoryStream ms = new MemoryStream(byteArrayIn);
                Image returnImage = Image.FromStream(ms);
                return returnImage;
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        public byte[] imageToByteArray(System.Drawing.Image imageIn)
        {
            try
            {
                MemoryStream ms = new MemoryStream();
                imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Gif);
                return ms.ToArray();
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        public String insertImage(string imageuri, string uname, string pwd)
        {
            object missing = Type.Missing;
            string message = "";

            try
            {
                byte[] bytearray = downloadData(imageuri, uname, pwd);
                Image img = byteArrayToImage(bytearray);

                PPT.Slide slide = (PPT.Slide)Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

                Clipboard.SetImage(img);
                slide.Shapes.Paste();
                Clipboard.Clear();
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }


            return message;
        }

        public String getPresentationName()
        {
            string filename = "";
            try{
                filename = Globals.ThisAddIn.Application.ActivePresentation.Name;
            }catch(Exception e)
            {
                string errorMsg = e.Message;
                filename = "error: " + errorMsg;
            }

            return filename;
        }

        public String getPresentationPath()
        {
            string path = "";
            try
            {
                path = Globals.ThisAddIn.Application.ActivePresentation.Path;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                path = "error: " + errorMsg;
            }
            return path;
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

        //need a param here for class type? right now works from filename for Word, Excel; but we could specify classname
        public string embedOLE(string path, string title, string url, string user, string pwd)
        {
            string message="";
            string tmpdoc = "";
            object missing = System.Type.Missing;
            bool proceed = false;
            int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
            float left=60;
            float top=105;
            float width=600;
            float height=300;

            if(title.EndsWith(".docx") || title.EndsWith(".docm") ||
               title.EndsWith(".dotx") || title.EndsWith(".dotm"))
            {
                left=220;
                width=300;
            }
                         try
                         {
                             tmpdoc = path + title;
                             downloadFile(url, tmpdoc, user, pwd);
                             proceed = true;

                         }
                         catch (Exception e)
                         {
                             string errorMsg = e.Message;
                             message = "error: " + errorMsg;
                         }

                         try
                         {
                             if (proceed)
                             {
                                 //defaulting args here.  these could be parameters.
                                 //you specify classtype or filename, not both
                                 Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes.AddOLEObject(left,top,width,height, "", tmpdoc, Microsoft.Office.Core.MsoTriState.msoFalse, "", 0, "", Microsoft.Office.Core.MsoTriState.msoFalse);
                             }
                         }
                         catch (Exception e)
                         {
                             string errorMsg = e.Message+e.StackTrace;
                             message = "error: " + errorMsg;
                         }

            return message;
        }

        public String openPPTX(string path, string title, string url, string user, string pwd)
        {
            //MessageBox.Show("in the addin path:"+path+  "      title:"+title+ "   uri: "+url+"user"+user+"pwd"+pwd);
            string message = "";
            object missing = Type.Missing;
            string tmpdoc = "";

            try
            {
                tmpdoc = path + title;
                downloadFile(url, tmpdoc, user, pwd);
                PPT.Presentation ppt = Globals.ThisAddIn.Application.Presentations.Open(tmpdoc, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue);
            }
            catch (Exception e)
            {
                //not always true, need to improve error handling or message or both
                string origmsg = "A presentation with the name '" + title + "' is already open. You cannot open two documents with the same name, even if the documents are in different \nfolders. To open the second document, either close the document that's currently open, or rename one of the documents.";
                MessageBox.Show(origmsg);
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string insertSlide(string tmpPath, string filename, string slideidx,string url,string user, string pwd, string retain)
        {

            string message = "";
            object missing = Type.Missing;
            string sourcefile = "";
            string path = getTempPath();
            bool retainformat = false;
            bool proceed = false;

            PPT.Slides ss = Globals.ThisAddIn.Application.ActivePresentation.Slides;

            if (retain.ToLower().Equals("true"))
                retainformat = true;

            try
            {
                sourcefile = path + filename;
                if (FileInUse(sourcefile))
                {
                    string origmsg = "A presentation with the name '" + filename + "' is already open. You cannot open two documents with the same name, even if the documents are in different \nfolders. To open the second document, either close the document that's currently open, or rename one of the documents.";
                    MessageBox.Show(origmsg);
                    
                }
                else
                {
                    downloadFile(url, sourcefile, user, pwd);
                    proceed = true;
                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            
            try
            {
                if (proceed)
                {
                    PPT.Presentation sourcePres = Globals.ThisAddIn.Application.Presentations.Open(sourcefile, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);
                    int num = Convert.ToInt32(slideidx);
                    copyPasteSlideToActive(sourcePres, num, retainformat);
                    sourcePres.Close();
                    sourcePres = null;
                }
            }
            catch(Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;    
            }

            return message;
        }

        public string copyPasteSlideToActive(PPT.Presentation sourcePres, int slideidx, bool retain)
        {
            string message = "";
            PPT.Presentation activePres = Globals.ThisAddIn.Application.ActivePresentation;
            PPT.Slides activeSlides = activePres.Slides;
            PPT.Slides sourceSlides = sourcePres.Slides;

            for (int x = 1; x <= sourceSlides.Count; x++)
            {
                int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
                int id = sourceSlides[x].SlideID;

                if (sourceSlides[x].SlideIndex == slideidx)
                {
                    sourceSlides.FindBySlideID(id).Copy();
                    try
                    {
                        if (retain)
                        {
                            activeSlides.Paste(sid).FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoTrue;
                            Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[sid].Select();
                            PPT.SlideRange sr = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange;
                            sr.Design = sourcePres.SlideMaster.Design;
                        }
                        else
                        {
                            activeSlides.Paste(sid).FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
                            Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[sid].Select();
                            PPT.SlideRange sr = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange;
                            sr.Design = Globals.ThisAddIn.Application.ActivePresentation.SlideMaster.Design;
                        }
                     }
                    catch (Exception e)
                    {
                        string errorMsg = e.Message;
                        message = "error: " + errorMsg; 
                    }
                }


            }

            return message;
        }

        public string convertFilenameToImageDir(string filename)
        {
            string imgDir = "";
            string tmpDir = "";
            string fname = "";

            string[] split = filename.Split(new Char[] { '\\' });
            fname = split.Last();
            tmpDir = filename.Replace(fname, "");
            fname = fname.Replace(".pptx", "_PNG");
            imgDir = fname; //getTempPath() + fname;
            return imgDir;

        }

        public string useSaveFileDialog()
        {
            string message = "";
            try
            {
                Prompt p = new Prompt();
                p.ShowDialog();
                string filename = p.pfilename;
                if (filename.Trim().Equals("") || filename.Trim() == null)
                {
                    //do nothing
                }
                else if (!filename.EndsWith(".pptx"))
                {
                        message = filename + ".pptx";
                }
                else
                {
                        message = filename;
                }
                
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg; 
            }

            return message;
        }

        private void downloadFile(string url, string sourcefile, string user, string pwd)
        {
            try
            {
                System.Net.WebClient Client = new System.Net.WebClient();
                Client.Credentials = new System.Net.NetworkCredential(user, pwd);
                Client.DownloadFile(url, sourcefile);
                Client.Dispose();
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        private void uploadData(string url, byte[] content, string user, string pwd)
        {
            try
            {
                System.Net.WebClient Client = new System.Net.WebClient();
                Client.Headers.Add("enctype", "multipart/form-data");
                Client.Headers.Add("Content-Type", "application/octet-stream");
                Client.Credentials = new System.Net.NetworkCredential(user, pwd);

                Client.UploadData(url, "POST", content);
                Client.Dispose();
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        private byte[] downloadData(string url, string user, string pwd)
        {
            byte[] bytearray;
            try
            {
                System.Net.WebClient Client = new System.Net.WebClient();
                Client.Credentials = new System.Net.NetworkCredential(user, pwd);
                bytearray = Client.DownloadData(url);
                Client.Dispose();
            }
            catch (Exception e)
            {
                throw (e);
            }
            return bytearray;
        }

        public string saveActivePresentation(string filename, string url, string user , string pwd)
        {
            string message = "";
            try
            {
                FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                int length = (int)fs.Length;
                byte[] content = new byte[length];
                fs.Read(content, 0, length);

                try
                {
                    uploadData(url, content, user, pwd);
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                fs.Dispose();
                fs.Close();

            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string saveLocalCopy(string filename)
        {
            string message = "";

            try
            {
                PPT.Presentation pptx = Globals.ThisAddIn.Application.ActivePresentation;
                pptx.SaveAs(filename, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Microsoft.Office.Core.MsoTriState.msoFalse);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string saveActivePresentationAndImages(string saveasdirectory, string saveasname, string url, string user, string pwd)
        {
           //MessageBox.Show("SAVE PRES WITH IMAGES:\nsaveasdirectory: " + saveasdirectory + "\nsaveasname: " + saveasname+"\nurl: "+url); 
           //can just append /dir/path/etc to url= for insert in other locations.
           //may want to have parameter later
           string message = "";

           try
           {
            string fullfilenamewithpath = "";
            string imgdirwithpath = "";
            string filename = "";

            fullfilenamewithpath = saveasdirectory + saveasname; // useSaveFileDialog()+".pptx";
            filename = fullfilenamewithpath.Split(new Char[] { '\\' }).Last();

            saveLocalCopy(fullfilenamewithpath);    
            imgdirwithpath = getTempPath() + convertFilenameToImageDir(fullfilenamewithpath);

            saveImages(imgdirwithpath, url, user, pwd);
            string fullurl = url + "/" + filename;
            saveActivePresentation(fullfilenamewithpath, fullurl, user, pwd);
          
           }
           catch (Exception e)
           {
               string errorMsg = e.Message;
               message = "error: " + errorMsg;
           }

            return message;
        }

        public string saveImages(string imgdirwithpath, string url, string user, string pwd)
        {
            string message = "";
            string imgdir = imgdirwithpath.Split(new Char[] { '\\' }).Last();
        
            imgdir = "/" + imgdir; 
            PPT.Presentation ppt = Globals.ThisAddIn.Application.ActivePresentation;

            try
            {
                if (Directory.Exists(imgdirwithpath))
                {
                    string[] files = Directory.GetFiles(imgdirwithpath);
                    foreach (string s in files)
                    {
                        File.Delete(s);
                    }
                    Directory.Delete(imgdirwithpath);
                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                return message;
            }

            try
            {
                ppt.SaveAs(imgdirwithpath, PPT.PpSaveAsFileType.ppSaveAsPNG, Office.MsoTriState.msoFalse);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                return message;
            }

            string[] imgfiles = Directory.GetFiles(imgdirwithpath);

            foreach (string i in imgfiles)
            {
                string fname = i.Split(new Char[] { '\\' }).Last();
                string fileuri = imgdir + "/" + fname;
                string fullurl = url + fileuri;

                try
                {
                   
                    FileStream fs = new FileStream(i, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    int length = (int)fs.Length;
                    byte[] content = new byte[length];
                    fs.Read(content, 0, length);

                    try
                    {
                        uploadData(fullurl, content,user,pwd);
                    }
                    catch (Exception e)
                    {
                        string errorMsg = e.Message;
                        message = "error: " + errorMsg;
                    }
                    
                    fs.Dispose();
                    fs.Close();
                   
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }
            }

            return message;
        }

        public string insertText(string txt)
        {
            string message = "";
            try
            {
                string orig =  Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange.Text;
                Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange.Text = orig + txt;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                if (e.Message.Contains("Nothing appropriate is currently selected"))
                {
                    MessageBox.Show("Please set text insertion point with cursor selection.");
                }

            }

            return message;
           
        }

        class MYTABLE
        {
            public List<string> headers { get; set;}
            public List<string[]> values { get; set; }
        }

        //this is interesting, not sure if we can use
        //ppt, word, and excel "tables" are all different
        //ultimately might want a server side transform to create generalized table,
        //and send  XML representation of tbl to function in Addin insertTable(string XML)
        public string insertJSONTable(string table) //parameterize rows, columns, vals
        {
            string message = "";
            //MessageBox.Show("table: " + table);
            //JavaScriptSerializer ser = new JavaScriptSerializer();
            try
            {
                object missing = System.Type.Missing;
                MYTABLE mytable = new MYTABLE();
                mytable = new JavaScriptSerializer().Deserialize<MYTABLE>(table);

                List<string> labels = mytable.headers;
                List<string[]> vals = mytable.values;

                //labels = 1 row, labels.Count = #columns
                //val count = rows, val count + 1 (for labels) = total # of rows
                int columnslength = labels.Count;
                int rowslength = vals.Count + 1;

                //create table
                int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
                PPT.Shape s = Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes.AddTable(rowslength, columnslength,50,50,450,70);
                PPT.Table tbl = s.Table;

                int lblcolidx = 1;
                foreach (string l in labels)
                {
                  PPT.Cell cell = tbl.Rows[1].Cells[lblcolidx];
                  cell.Shape.TextFrame.TextRange.Text = l;
                  lblcolidx++;
                
                }

                int rowidx = 2;
                foreach (string[] v in vals)
                {
                  //MessageBox.Show("Starting loop");
                  int colidx = 1;
                  string[] vs = v;
                  for(int i=0;i<vs.Length;i++)
                  {
                      //MessageBox.Show("Value"+vs[i]);
                      PPT.Cell cell = tbl.Rows[rowidx].Cells[colidx];
                      cell.Shape.TextFrame.TextRange.Text = vs[i];
                      colidx++;
                  }
                  rowidx++;
                  //MessageBox.Show("vs count "+v.Length);
                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                MessageBox.Show("ERROR: " + e.Message + "      " + e.StackTrace);
            }
           
            return message;
        }

    }
}
