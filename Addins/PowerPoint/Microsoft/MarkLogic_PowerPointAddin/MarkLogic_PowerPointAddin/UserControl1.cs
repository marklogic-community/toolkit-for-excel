/*Copyright 2009-2010 Mark Logic Corporation

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
//using PwrPt = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.IO;
//using DocumentFormat.OpenXml.Packaging; //OpenXML sdk
using Office = Microsoft.Office.Core;
using Microsoft.Win32;
using PPT = Microsoft.Office.Interop.PowerPoint;
//using OX = DocumentFormat.OpenXml.Packaging;
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
        //private string addinVersion = "@MAJOR_VERSION.@MINOR_VERSION@PATCH_VERSION";
        private string addinVersion = "1.1-1";
        HtmlDocument htmlDoc;

        public PPT.ApplicationClass  ppta = null;
        public bool firePptCloseEvent = true;

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


                if (ac.getEventsEnabled())
                {
                    //Event Handling in TKEvents.cs
                    ppta = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
                    ppta.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                    System.Runtime.InteropServices.ComTypes.IConnectionPoint mConnectionPoint;
                    System.Runtime.InteropServices.ComTypes.IConnectionPointContainer cpContainer;
                    int mCookie;

                    cpContainer =
                    (System.Runtime.InteropServices.ComTypes.IConnectionPointContainer)ppta;
                    Guid guid = typeof(Microsoft.Office.Interop.PowerPoint.EApplication).GUID;
                    cpContainer.FindConnectionPoint(ref guid, out mConnectionPoint);
                    mConnectionPoint.Advise(this, out mCookie);
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
                PPT.Presentation pres = Globals.ThisAddIn.Application.ActivePresentation;
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
                PPT.Presentation pres = Globals.ThisAddIn.Application.ActivePresentation;
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
                PPT.Presentation pres = Globals.ThisAddIn.Application.ActivePresentation;
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
                PPT.Presentation pres = Globals.ThisAddIn.Application.ActivePresentation;
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

        public String insertImage(string imageuri, string uname, string pwd)
        {
            object missing = Type.Missing;
            string message = "";

            try
            {
                byte[] bytearray = TKUtilities.downloadData(imageuri, uname, pwd);
                Image img = TKUtilities.byteArrayToImage(bytearray);

                PPT.Slide slide = (PPT.Slide)Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

                Clipboard.SetImage(img);
                PPT.ShapeRange sr  = slide.Shapes.Paste();
                sr.Select(Microsoft.Office.Core.MsoTriState.msoFalse);
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
                             TKUtilities.downloadFile(url, tmpdoc, user, pwd);
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
                                 Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes.AddOLEObject(left, top, width, height, "", tmpdoc, Microsoft.Office.Core.MsoTriState.msoFalse, "", 0, "", Microsoft.Office.Core.MsoTriState.msoFalse);

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
            string message = "";
            object missing = Type.Missing;
            string tmpdoc = "";

            try
            {
                tmpdoc = path + title;
                TKUtilities.downloadFile(url, tmpdoc, user, pwd);
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
            string path = tmpPath; ///changed from getTempPath()
            bool retainformat = false;
            bool proceed = false;

            PPT.Slides ss = Globals.ThisAddIn.Application.ActivePresentation.Slides;

            if (retain.ToLower().Equals("true"))
                retainformat = true;

            try
            {
                sourcefile = path + filename;
                if (TKUtilities.FileInUse(sourcefile))
                {
                    string origmsg = "A presentation with the name '" + filename + "' is already open. You cannot open two documents with the same name, even if the documents are in different \nfolders. To open the second document, either close the document that's currently open, or rename one of the documents.";
                    MessageBox.Show(origmsg);
                    
                }
                else
                {
                    TKUtilities.downloadFile(url, sourcefile, user, pwd);
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
                    int num = TKUtilities.getInt32FromString(slideidx);
                    copyPasteSlideToActive(sourcePres, num, retainformat);
                    firePptCloseEvent = false;
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
                        message = "error: COPYPASTE" + errorMsg; 
                    }
                }


            }

            return message;
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
                    TKUtilities.uploadData(url, content, user, pwd);
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
            //MessageBox.Show("filename" + fullfilenamewithpath);
            saveLocalCopy(fullfilenamewithpath);    
            imgdirwithpath = getTempPath() + TKUtilities.convertFilenameToImageDir(fullfilenamewithpath);

            //MessageBox.Show("imgdirwithpath"+imgdirwithpath);
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
                //MessageBox.Show("imgdir with path before save"+imgdirwithpath);
                ppt.SaveAs(imgdirwithpath, PPT.PpSaveAsFileType.ppSaveAsPNG, Office.MsoTriState.msoFalse);
                //MessageBox.Show("imgdir with path after save" + imgdirwithpath);
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
                        TKUtilities.uploadData(fullurl, content,user,pwd);
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

        //ppt, word, and excel (let alone html, etc.) "tables" are all different
        //ultimately might want a server side transform to create generalized table,
        //and send XML representation of tbl to function in Addin insertTable(string XML)
        public string insertJSONTable(string table) //parameterize rows, columns, vals
        {
            string message = "";

            try
            {
                object missing = System.Type.Missing;
                JsonTable mytable = new JsonTable();
                mytable = new JavaScriptSerializer().Deserialize<JsonTable>(table);

                List<string> labels = mytable.headers;
                List<string[]> vals = mytable.values;

                //labels = 1 row, labels.Count = #columns
                //val count = rows, val count + 1 (for labels) = total # of rows
                int columnslength = labels.Count;
                int rowslength = vals.Count + 1;

                int tmpwidth = 100 * columnslength;
                int tmpheight = 30 * rowslength;
                int width = (tmpwidth > 600)? 600: tmpwidth;
                int height = (tmpheight > 600) ? 450 : tmpheight;

                //create table
                int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
                PPT.Shape s = Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes.AddTable(rowslength, columnslength,60,50,width,height);
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

                  int colidx = 1;
                  string[] vs = v;
                  for(int i=0;i<vs.Length;i++)
                  {
                      PPT.Cell cell = tbl.Rows[rowidx].Cells[colidx];
                      cell.Shape.TextFrame.TextRange.Text = vs[i];
                      colidx++;
                  }
                  rowidx++;
                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                //MessageBox.Show("ERROR: " + e.Message + "      " + e.StackTrace);
            }
           
            return message;
        }
//BEGIN TK 1.1
        public string getSlideName()
        {
            string message = "";

            try
            {
                PPT.Slide slide = (PPT.Slide)Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                string slideName = slide.Name;
                message = slideName.ToString();
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string getSlideIndex()
        {
            string message = "";

            try
            {
                PPT.Slide slide = (PPT.Slide)Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
                int slideIndex = slide.SlideIndex;
               
                message = slideIndex.ToString();
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string getPresentationSlideCount()
        {
            string message = "";

            try
            {
                int slideCount = Globals.ThisAddIn.Application.ActivePresentation.Slides.Count;
                message = slideCount.ToString();
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }
           
            return message;
        }

        public string addSlideTag(string slideIndex, string tagName, string tagValue)
        {

            string message = "";
            try
            {
                int index = TKUtilities.getInt32FromString(slideIndex);
                Globals.ThisAddIn.Application.ActivePresentation.Slides[index].Tags.Add(tagName, tagValue);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string deleteSlideTag(string slideIndex, string tagName)
        {
            string message = "";
            try
            {
                int index = TKUtilities.getInt32FromString(slideIndex);
                Globals.ThisAddIn.Application.ActivePresentation.Slides[index].Tags.Delete(tagName);
          
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string getSlideTags(string slideIndex)
        {
            string message = "";
            int index = TKUtilities.getInt32FromString(slideIndex);
            string jsonSlideTags = "{\"tags\":[";
            PPT.Tags sTags = Globals.ThisAddIn.Application.ActivePresentation.Slides[index].Tags;

            try
            {
                for (int j = 1; j <= sTags.Count; j++)
                {
                    jsonSlideTags += "{\"name\":\"" + sTags.Name(j) +
                                  "\",\"value\":\"" + sTags.Value(j) +
                                  "\"},";

                }

                if (jsonSlideTags.EndsWith(","))
                    jsonSlideTags = jsonSlideTags.Substring(0, jsonSlideTags.Length - 1);

                jsonSlideTags += "]}";
                message = jsonSlideTags;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }


        public string getShapesTest()
        {
            string message = "";

            //float height = 115;
            PPT.Shapes s = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.Shapes;
            MessageBox.Show("Shapes Count: " + s.Count);
            foreach (PPT.Shape sp in s)
            {
                MessageBox.Show("FORMAT: " + sp.Type +
                                "FONT NAME: "+ sp.TextFrame.TextRange.Font.Name+
                                "FONT SIZE: "+sp.TextFrame.TextRange.Font.Size +
                                "FONT COLOR: "+sp.TextFrame.TextRange.Font.Color.RGB +
                                "CENTERED: "+sp.TextFrame.TextRange.ParagraphFormat.Alignment);
            }

          //s.AddShape(Microsoft.Office.Core.MsoShapeType.msoPlaceholder, 0, 0, 100, 100);
          //  PPT.Shape me = s.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 612, height);
          //  me.TextFrame.TextRange.Font.Name = "Calibri";
          //  me.TextFrame.TextRange.Font.Size = 44;
          //  me.TextFrame.TextRange.Text = "My Title";
          //s.AddPlaceholder(PPT.PpPlaceholderType.ppPlaceholderSubtitle,1,1,100,100 );

            return message;
        }

        public string addShapeTag(string slideIndex, string shapeName, string tagName, string tagValue)
        {

            string message = "";
            try
            {
                int idx = Convert.ToInt32(slideIndex);
                Globals.ThisAddIn.Application.ActivePresentation.Slides[idx].Shapes[shapeName].Tags.Add(tagName, tagValue) ;
                //Globals.ThisAddIn.Application.ActivePresentation.Windows.Application.ActiveWindow.Selection.SlideRange.Shapes[shapeName].Tags.Add(tagName, tagValue);
                
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string deleteShapeTag(string slideIndex, string shapeName, string tagName)
        {
            string message = "";
            int idx = Convert.ToInt32(slideIndex);
             
            try
            {
                Globals.ThisAddIn.Application.ActivePresentation.Slides[idx].Shapes[shapeName].Tags.Delete(tagName);
                //Globals.ThisAddIn.Application.ActivePresentation.Windows.Application.ActiveWindow.Selection.SlideRange.Shapes[shapeName].Tags.Delete(tagName);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        //do we want this for slide too? probably.
        public string getShapeRangeName()
        {
            string message = "";
            try
            {
                PPT.ShapeRange sr = Globals.ThisAddIn.Application.ActivePresentation.Windows.Application.ActiveWindow.Selection.ShapeRange;
                message = sr.Name;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string setShapeRangeName(string slideIndex, string shapeName, string newShapeName)
        {
            string message = "";
            int idx = Convert.ToInt32(slideIndex);
            try{
                //Though a ShapeRange can contain multiple shapes, 
                //to set the Name property, A ShapeRange can only contain 1 Shape.
                //developers should use getShapeRangeCount()and insure it = 1 to validate before calling this function
               // PPT.ShapeRange sr = Globals.ThisAddIn.Application.ActivePresentation.Windows.Application.ActiveWindow.Selection.ShapeRange;
               // sr.Name = newshapename;

                //PPT.Selection sel = Globals.ThisAddIn.Application.ActivePresentation.Windows.Application.ActiveWindow.Selection;
                //PPT.Shape s = sel.SlideRange.Shapes[shapeName];
                PPT.Shape s = Globals.ThisAddIn.Application.ActivePresentation.Slides[idx].Shapes[shapeName];
                s.Name = newShapeName;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;

        }
      
        public string getSlideShapeNames(string slideIndex)
        {
            string message = "";
            int index = TKUtilities.getInt32FromString(slideIndex);
            try
            {
                PPT.Slide slide = Globals.ThisAddIn.Application.ActivePresentation.Slides[index];
                PPT.Shapes slideShapes = slide.Shapes;
                //PPT.Shapes slideShapes = Globals.ThisAddIn.Application.ActivePresentation.Windows.Application.ActiveWindow.Selection.SlideRange.Shapes;
                foreach (PPT.Shape s in slideShapes)
                {
                    message = message + s.Name + "|";
                }

                message = message.Substring(0, message.Length - 1);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }
        //gets currently selected shape range names
        //remember, you can select multiple
        //but when using the api, you can only operate on one shape at a time
        public string getShapeRangeShapeNames()
        {
            string message = "";
            try
            {
                PPT.ShapeRange sr = Globals.ThisAddIn.Application.ActivePresentation.Windows.Application.ActiveWindow.Selection.ShapeRange;

                for (int i = 1; i <= sr.Count; i++)
                {
                    PPT.Shape s = sr[i];
                    message = message + s.Name + "|";
                }

                message = message.Substring(0, message.Length - 1);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        //operates on currently selected
        //can be more than one
//Do we want similar function that adds tag by slideindex and shapename?
        public string addShapeRangeTag(string tagName, string tagValue)
        {

            string message = "";
            try
            {   
                //int idx = Convert.ToInt32(shapeid);
                PPT.ShapeRange sr = Globals.ThisAddIn.Application.ActivePresentation.Windows.Application.ActiveWindow.Selection.ShapeRange;

                for (int i = 1; i <= sr.Count; i++)
                {
                    PPT.Shape s = sr[i];
                    s.Tags.Add(tagName, tagValue);
                 }
                
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string addPresentationTag(string tagName, string tagValue)
        {
            string message = "";
            try
            {   //BS PROOF
                //You can't add a Tag with a name that already exists
                //no indication, the function is truly void
                //so check before adding using (getPresentationTags() in the .js)
                Globals.ThisAddIn.Application.ActivePresentation.Tags.Add(tagName, tagValue);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string deletePresentationTag(string tagName)
        {
            string message = "";
            try
            {
                Globals.ThisAddIn.Application.ActivePresentation.Tags.Delete(tagName);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string getPresentationTags()
        {
            string message = "";
            string jsonPresTags = "{\"tags\":[";
            PPT.Tags pTags = Globals.ThisAddIn.Application.ActivePresentation.Tags;

            try
            {
                for (int j = 1; j <= pTags.Count; j++)
                {
                    jsonPresTags += "{\"name\":\"" + pTags.Name(j) +
                                  "\",\"value\":\"" + pTags.Value(j) +
                                  "\"},";

                }

                if (jsonPresTags.EndsWith(","))
                    jsonPresTags = jsonPresTags.Substring(0, jsonPresTags.Length - 1);

                jsonPresTags += "]}";
                message=jsonPresTags;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }
          
            return message;
        }

        //need to find out about embedded shapes? is there such a thing?
        public string getShapeRangeCount()
        {
            string message = "";
            try
            {
                int count = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Count;
                message = count.ToString();
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string getShapeRangeView(string slideIndex, string shapeName)
        {
            string message = "";
            int index = TKUtilities.getInt32FromString(slideIndex);

            try
            {
                //PPT.Selection sel = Globals.ThisAddIn.Application.ActivePresentation.Windows.Application.ActiveWindow.Selection;
                //PPT.Shape s = sel.SlideRange.Shapes[shapeName];

                PPT.Shape s = Globals.ThisAddIn.Application.ActivePresentation.Slides[index].Shapes[shapeName];

                PPT.Tags sTags = s.Tags;
                string jsonRange = "";
                object t = Office.MsoTriState.msoTrue;


                //BASIC SHAPE
                jsonRange = //"{\"shapeRange\":" +
                            "{\"name\":\"" + s.Name + "\"," +
                            "\"left\":\"" + s.Left + "\"," +
                            "\"top\":\"" + s.Top + "\"," +
                            "\"height\":\"" + s.Height + "\"," +
                            "\"width\":\"" + s.Width + "\",";


                //DO CHECK ON TYPES AND UPDATE ACCORDINGLY, CAN'T ROUNDTRIP PLACEHOLDERS
                //PPT OBJECT MODEL ONLY ALLOWS DELETE/ADD OF EXISTING, CAN'T CREATE NEW PLACEHOLDER
                //TEXTBOX SHAPE
                //EXTRA COMMA
                if (s.Type.Equals(Office.MsoShapeType.msoPlaceholder) && s.HasTextFrame.Equals(t))
                {
                    jsonRange += "\"type\":\"" + Office.MsoShapeType.msoTextBox + "\"";
                }
                else
                {
                    jsonRange += "\"type\":\"" + s.Type + "\"";
                }

          /*    FOLLOWING REPLACED WITH PARAGRAPH ARRAY BELOW 
                if (s.HasTextFrame.Equals(t))
                {   
                    //CAN'T RETURN TEXT AS STRING IS JACKED UP WITH MULTIPLE PARAS/RUNS
                    jsonRange += //"\"text\":\"" + s.TextFrame.TextRange.Text + "\"," +
                                  "\"text\":\"" + "FUBAR" + "\"," +
                                 "\"fontName\":\"" + s.TextFrame.TextRange.Font.Name + "\"," +
                                 "\"fontSize\":\"" + s.TextFrame.TextRange.Font.Size + "\"," +
                                 "\"fontRGB\":\"" + s.TextFrame.TextRange.Font.Color.RGB + "\"," +
                                // "\"paragraphAlignment\":\"" + s.TextFrame.TextRange.ParagraphFormat.Alignment + "\"," +
                                 "\"textOrientation\":\"" + s.TextFrame.Orientation + "\",";
                }
           */
                //PARAGRAPHS WITHIN TEXTBOX
                //PARAGRAPHS HAVE ALIGNMENT
                //PARAGRAPHGS CONTAIN RUNS
                //RUNS HAVE STYLES
                if (s.HasTextFrame.Equals(t))
                {
                    //After round trip, may add these. 
                    //This way, if you have no text (no paras, no runs), you know
                    //the style of the object you are potentially inserting into
                    //"\"pargraphFontName\":\"" + s.TextFrame.TextRange.Font.Name + "\"," +
                    //"\"paragraphFontSize\":\"" + s.TextFrame.TextRange.Font.Size + "\"," +
                    //"\"paragraphFontRGB\":\"" + s.TextFrame.TextRange.Font.Color.RGB + "\"," +
                    jsonRange += ",\"textOrientation\":\"" + s.TextFrame.Orientation + "\",";
                    jsonRange += "\"paragraphs\":[";
                    try
                    {
                        //1000 as when length is greater than number of paras in range, 
                        //returns all paras to range.paras.count 
                        PPT.TextRange allParas = s.TextFrame.TextRange.Paragraphs(1, 1000);
                  
                        foreach (PPT.TextRange para in allParas)
                        {
                            jsonRange += "{\"paragraphAlignment\":\"" + para.ParagraphFormat.Alignment + "\",";
                            jsonRange += "\"paragraphBulletType\":\"" + para.ParagraphFormat.Bullet.Type.ToString().Normalize().Trim() + "\",";

                           
                            //needs to be its own object, may require special massage for paragraphs
                            //para.ParagraphFormat.Bullet.*
                            jsonRange += "\"runs\":[";
                            
                            foreach (PPT.TextRange run in para.Runs(1, 1000))
                            {
                                //Bulleted Lists and Multiple Paras add end of lines and other cruft
                                //IE will choke on with eval of JSON
                                //Still needs work
                                string text = run.Text.Normalize().Trim();

                                jsonRange += "{\"fontName\":\"" + run.Font.Name + "\"," +
                                              "\"fontSize\":\"" + run.Font.Size + "\"," +
                                               "\"fontRGB\":\"" + run.Font.Color.RGB + "\"," +
                                               "\"fontItalic\":\"" + run.Font.Italic + "\"," +
                                               "\"fontUnderline\":\"" + run.Font.Underline + "\"," +
                                              
                                               "\"fontBold\":\"" + run.Font.Bold + "\"," +
                                                "\"text\":\"" + text +
                                                "\"},";
                             

                                /*
                                MessageBox.Show("font: " + run.Font.Name +
                                                " rgb: " + run.Font.Color.RGB +
                                                " size: " + run.Font.Size +
                                                " text: " + run.Text);
                                */
                            }

                            if (jsonRange.EndsWith(","))
                                jsonRange = jsonRange.Substring(0, jsonRange.Length - 1);

                            jsonRange += "]";  //end of run array
                            jsonRange += "},"; //next paragraph
                        }

                        if (jsonRange.EndsWith(","))
                            jsonRange = jsonRange.Substring(0, jsonRange.Length - 1);

                        jsonRange += "]"; //end of paragraph array

                    }
                    catch (Exception e)
                    {
                        //MessageBox.Show("PARAGRAPHS ERROR" + e.Message);
                        string errorMsg = e.Message;
                        message = "error: " + errorMsg;
                    }
                }

                if (s.Type.Equals(Office.MsoShapeType.msoPicture))
                {
                    jsonRange+=",\"pictureFormat\":" +
                                  "{\"brightness\":\"" + s.PictureFormat.Brightness + "\"," +
                                   "\"colorType\":\"" + s.PictureFormat.ColorType + "\"," +
                                   "\"contrast\":\"" + s.PictureFormat.Contrast + "\"," +
                                   "\"cropBottom\":\"" + s.PictureFormat.CropBottom + "\"," +
                                   "\"cropLeft\":\"" + s.PictureFormat.CropLeft + "\"," +
                                   "\"cropRight\":\"" + s.PictureFormat.CropRight + "\"," +
                                   "\"cropTop\":\"" + s.PictureFormat.CropTop + "\"," +
                                   "\"transparencyColor\":\"" + s.PictureFormat.TransparencyColor + "\"," +
                                   "\"transparentBackground\":\"" + s.PictureFormat.TransparentBackground + "\"}";
                    //MessageBox.Show(s.PictureFormat.Brightness+"");
                }


            /*  PPT.TextRange paraTest = s.TextFrame.TextRange.Paragraphs(1,1000);
                MessageBox.Show("# paragraphs: "+paraTest.Count);

               
                foreach (PPT.TextRange para in paraTest)
                {
                    MessageBox.Show("Alignment: " + para.ParagraphFormat.Alignment);
                   
                    foreach (PPT.TextRange run in para.Runs(1,1000))
                    {
                       
                        MessageBox.Show("font: " + run.Font.Name +
                                        " rgb: " + run.Font.Color.RGB +
                                        " size: " + run.Font.Size +
                                        " text: " + run.Text);
                                       
                    }

                }
                */
                                 
                jsonRange += ",\"tags\":[";

                try
                {
                    for (int j = 1; j <= sTags.Count; j++)
                    { 
                       jsonRange += "{\"name\":\"" + s.Tags.Name(j) +
                                     "\",\"value\":\"" + sTags.Value(j) +
                                     "\"},";

                    }

                    if(jsonRange.EndsWith(","))
                       jsonRange = jsonRange.Substring(0, jsonRange.Length - 1);

                       jsonRange += "]";
                }
                catch (Exception e)
                {
                    //MessageBox.Show("TAGS ERROR" + e.Message);
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                }

                jsonRange += "}";
                         //}"; end of ShapeRange

                message = jsonRange; 

            }
            catch (Exception e)
            {
                //string donothing_removewarning = e.Message;
                //MessageBox.Show("getShapeRangeInfoError: " + e.Message);
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            //MessageBox.Show(message);
            return message;
        }

        public string addShapeTags(string slideIndex, string shapeName, string jsonTags)
        {
            string message = "";
            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                List<TagView> shapeTags;
                shapeTags = serializer.Deserialize<List<TagView>>(jsonTags);

                int idx = TKUtilities.getInt32FromString(slideIndex);

                PPT.Shape shapeToTag = Globals.ThisAddIn.Application.ActivePresentation.Slides[idx].Shapes[shapeName];


                for (int i = 0; i < shapeTags.Count; i++)
                {
                    TagView t = new TagView();
                    t = shapeTags[i];

                    if (shapeToTag != null)
                        shapeToTag.Tags.Add(t.name, t.value);

                }

            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }
            return message;
        }

        public string addSlideTags(string slideIndex, string jsonTags)
        {
            string message = "";
            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                List<TagView> shapeTags;
                shapeTags = serializer.Deserialize<List<TagView>>(jsonTags);

                int idx = TKUtilities.getInt32FromString(slideIndex);

                PPT.Slide slideToTag = Globals.ThisAddIn.Application.ActivePresentation.Slides[idx];

                for (int i = 0; i < shapeTags.Count; i++)
                {
                    TagView t = new TagView();
                    t = shapeTags[i];

                    if (slideToTag != null)
                        slideToTag.Tags.Add(t.name, t.value);
                }

            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }
            return message;
        }

        public string addPresentationTags(string jsonTags)
        {
            string message = "";
            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                List<TagView> shapeTags;
                shapeTags = serializer.Deserialize<List<TagView>>(jsonTags);

                PPT.Presentation presToTag = Globals.ThisAddIn.Application.ActivePresentation;

                for (int i = 0; i < shapeTags.Count; i++)
                {
                    TagView t = new TagView();
                    t = shapeTags[i];

                    if(presToTag!=null)
                      presToTag.Tags.Add(t.name, t.value);
                }

            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }
            return message;
        }


        public string addShape(string slideIndex, string jsonShape, string jsonTags, string jsonParas)
        {
            string message = "";
            PPT.Shape newShape = null;

            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();

                //Deserialize
                ShapeRangeView mysr = new ShapeRangeView();
                List<TagView> shapeTags;
                ParagraphView paragraphs = new ParagraphView();

                mysr = serializer.Deserialize<ShapeRangeView>(jsonShape);
                shapeTags = serializer.Deserialize<List<TagView>>(jsonTags);
                paragraphs = new JavaScriptSerializer().Deserialize<ParagraphView>(jsonParas);

                List<string> pAlignments = new List<string>();
                pAlignments = paragraphs.paragraphAlignment;

                List<string> pBullets = new List<string>();
                pBullets = paragraphs.paragraphBulletType;

                List<string[]> pRuns = new List<string[]>();
                pRuns = paragraphs.runs;


                int paragraphCount = pAlignments.Count;
                int runsCount = pRuns.Count;
                //MessageBox.Show("Para Count: " + paragraphCount);
                //MessageBox.Show("Run Count:  " + runsCount);

                /*
                MessageBox.Show("Type: "+mysr.type+"\n"+
                                " Name: " + mysr.name + "\n" +
                                " Left: " + mysr.left + "\n" +
                                " Top: " + mysr.top + "\n" +
                                " Height: " + mysr.top + "\n" +
                                " Width: " + mysr.top + "\n" +
                                " Orientation: " + mysr.textOrientation + "\n" 
                               // " Text: "+mysr.text 
                                );
                */

                int idx = TKUtilities.getInt32FromString(slideIndex);
                PPT.Shapes slideShapes = Globals.ThisAddIn.Application.ActivePresentation.Slides[idx].Shapes;

                if(mysr.type.Equals("msoTextBox"))
                {

                 Office.MsoTextOrientation  addOrientation;
                 addOrientation = TKUtilities.getTextOrientation(mysr.textOrientation);

                 float addLeft;
                 float addTop;
                 float addWidth;
                 float addHeight;
 
                 addLeft = TKUtilities.getFloatFromString(mysr.left);
                 addTop = TKUtilities.getFloatFromString(mysr.top);
                 addWidth = TKUtilities.getFloatFromString(mysr.width);
                 addHeight = TKUtilities.getFloatFromString(mysr.height);

                 newShape = slideShapes.AddTextbox(addOrientation,addLeft,addTop,addWidth,addHeight);
                 message = newShape.Name;
                 
                 //need to get Paras in, then loop through to add paras, runs, etc.
                 //using insert after

                 for (int i = 0; i < paragraphCount; i++)
                 {
                     //newShape.TextFrame.TextRange.ParagraphFormat.Alignment = TKUtilities.getParagraphAlignment(pAlignments[i]);
                     //newShape.TextFrame.TextRange.ParagraphFormat.Bullet.Type = TKUtilities.getParagraphBulletType(pBullets[i]);
                     newShape.TextFrame.TextRange.InsertAfter(" ");

                     PPT.TextRange pRange = newShape.TextFrame.TextRange.Paragraphs(i + 1, i + 1);
                     //Setting explicitly and defining in runs
                     //it's possible to define it a run, and have it set to paragraph
                     pRange.ParagraphFormat.Alignment = TKUtilities.getParagraphAlignment(pAlignments[i]);
                     pRange.ParagraphFormat.Bullet.Type = TKUtilities.getParagraphBulletType(pBullets[i]);
                    

                    
                     pRange.Font.Italic = Microsoft.Office.Core.MsoTriState.msoFalse;
                     pRange.Font.Underline = Microsoft.Office.Core.MsoTriState.msoFalse;
                     pRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoFalse;
                    

                     PPT.TextRange rRange;


                     for (int j = 0; j < pRuns.Count; j++ )
                     {
                         string[] finalRun = pRuns[j];
                         string runIdx = finalRun[0];
                         if (TKUtilities.getInt32FromString(runIdx) == i)
                         {
                             rRange = pRange.Runs(j + 1, j + 1);

                             string fontName = finalRun[1];
                             string fontSize = finalRun[2];
                             string fontRGB = finalRun[3];
                             string fontItalic = finalRun[4];
                             string fontUnderline = finalRun[5];
                             string fontBold = finalRun[6];
                             string text = finalRun[7];

                             if (text.Equals("") || text == null)
                                 text = " ";

                            // if (fontItalic.Equals("msoTrue"))
                             //    rRange.Font.Italic = Microsoft.Office.Core.MsoTriState.msoTrue;

                             rRange.InsertAfter(text);


                             //need to pass a continue on serialization then check here
                             //currently a space will be added for every run
                             if (i != paragraphCount)
                                 rRange.InsertAfter(" ");

                           
                             //MessageBox.Show("TEXT"+text);
                             rRange.Font.Name = fontName;
                             rRange.Font.Size = TKUtilities.getFloatFromString(fontSize);
                             rRange.Font.Color.RGB = TKUtilities.getInt32FromString(fontRGB);

                             
//serialize fontRGB for paragraph as well
                             if (!(pBullets[i].Equals("ppBulletNone")))
                             {
                                 pRange.Font.Color.RGB = TKUtilities.getInt32FromString(fontRGB);
                             }
                             //pRange.Font.Color.RGB = TKUtilities.getInt32FromString(fontRGB);

                             if (fontItalic.Equals("msoTrue"))
                             {
                                 rRange.Font.Italic = Microsoft.Office.Core.MsoTriState.msoTrue;
                             }

                             if (fontUnderline.Equals("msoTrue")){
                                 rRange.Font.Underline = Microsoft.Office.Core.MsoTriState.msoTrue;
                             }

                             if(fontBold.Equals("msoTrue")){
                                 rRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                             }

                             
      
                             /*
                             MessageBox.Show("MY RUN" + "\n" +
                                             "index: " + runIdx + "\n" +
                                             "fontName: " + fontName + "\n" +
                                             "fontSize: " + fontSize + "\n" +
                                             "fontRGB: " + fontRGB + "\n" +
                                             "fontItalic: " + fontItalic + "\n" +
                                             "fontUnderline: " + fontUnderline + "\n" +
                                             "text: " + text + "\n"
                                             );
                             */
                          
                         }
                         else
                         {
                             //do nothing
                             //unfortunately, have to do it this way
                             //until we can figure out the mysterious JavaScriptSerializer
                         }

                     }

                     if (i + 1  < paragraphCount)
                     {
                         //this followed by insertAfter(" ") adds new paragraph
                         newShape.TextFrame.TextRange.InsertAfter(Environment.NewLine);

                     }
                 }

                }

                for (int i = 0; i < shapeTags.Count; i++)
                {
                    TagView t = new TagView();
                    t = shapeTags[i];
                    if(newShape!=null)
                    newShape.Tags.Add(t.name,t.value);
                }

            }
            catch (Exception e)
            {
                //MessageBox.Show("addShapeError: " + e.Message +e.StackTrace);
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string setPictureFormat(string slideIndex, string shapeName, string jsonPictureFormat)
        {
            string message = "";
            try{
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                PictureFormatView picFormat;
                picFormat = serializer.Deserialize<PictureFormatView>(jsonPictureFormat);

                int idx = TKUtilities.getInt32FromString(slideIndex);

                PPT.Shape shapeToFormat = Globals.ThisAddIn.Application.ActivePresentation.Slides[idx].Shapes[shapeName];

                if (shapeToFormat.Type.Equals(Office.MsoShapeType.msoPicture))
                {

                    shapeToFormat.PictureFormat.Brightness = TKUtilities.getFloatFromString(picFormat.brightness);
                    shapeToFormat.PictureFormat.ColorType = TKUtilities.getColorType(picFormat.colorType);
                    shapeToFormat.PictureFormat.Contrast = TKUtilities.getFloatFromString(picFormat.contrast);
                    shapeToFormat.PictureFormat.CropBottom = TKUtilities.getFloatFromString(picFormat.cropBottom);
                    shapeToFormat.PictureFormat.CropLeft = TKUtilities.getFloatFromString(picFormat.cropLeft);
                    shapeToFormat.PictureFormat.CropRight = TKUtilities.getFloatFromString(picFormat.cropRight);
                    shapeToFormat.PictureFormat.CropTop = TKUtilities.getFloatFromString(picFormat.cropTop);

                    if (picFormat.transparencyBackground != null)
                    {
                        shapeToFormat.PictureFormat.TransparentBackground = TKUtilities.getTriState(picFormat.transparencyBackground);
                        shapeToFormat.PictureFormat.TransparencyColor = TKUtilities.getInt32FromString(picFormat.transparencyColor);
                    }
               
                }
                else
                {
                    message = "error: Can not apply picture format to shape that is not an image.";
                }
               
            }
            catch (Exception e)
            {
                //MessageBox.Show("setPictureFormatError: " + e.Message +e.StackTrace);
                string errorMsg = e.Message;
                message = message + "error: " + errorMsg;
            }
            return message;
        }

        public string deleteSlide(string slideIndex)
        {
            string message = "";
            int index = TKUtilities.getInt32FromString(slideIndex);
            try
            {
                PPT.Slide slideToDelete = Globals.ThisAddIn.Application.ActivePresentation.Slides[index];
                slideToDelete.Delete();
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string addSlide(string slideIndex, string customLayout)
        {
            string message = "";
            int idx = TKUtilities.getInt32FromString(slideIndex);
            PPT.PpSlideLayout layout = TKUtilities.getSlideLayout(customLayout);

            try
            {
                PPT.Slide slideToAdd = Globals.ThisAddIn.Application.ActivePresentation.Slides.Add(idx,layout);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }
        

        public string deleteShape(string slideIndex, string shapeName)
        {
            string message = "";
            int index = TKUtilities.getInt32FromString(slideIndex);

            try
            {
                PPT.Shape shapeToDelete = Globals.ThisAddIn.Application.ActivePresentation.Slides[index].Shapes[shapeName];
                shapeToDelete.Delete();
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
