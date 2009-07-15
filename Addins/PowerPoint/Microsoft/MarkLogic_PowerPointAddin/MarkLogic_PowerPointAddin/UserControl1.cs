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
                //MessageBox.Show("Unable to find configuration info. Please insure OfficeProperties.txt exists in your system temp directory.  If problems persist, please contact your system administrator.");
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
        /*
        private bool checkUrlInRegistry()
        {
            RegistryKey regKey1 = Registry.CurrentUser;
            regKey1 = regKey1.OpenSubKey(@"MarkLogicAddinConfiguration\PowerPoint");
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
       */
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
                        ids += c.Id + " ";// "U+016000";


                    }
                }

                char[] space = { ' ' };
                ids = ids.TrimEnd(space);

                //char[] tengwar = { 'U', '+', '0', '1', '6', '0', '0', '0' };
                //ids = ids.TrimEnd(tengwar);
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

        public String getFileName()
        {
            string filename = "";
            filename = Globals.ThisAddIn.Application.ActivePresentation.Name;
            return filename;
        }

        public String getPath()
        {
            string path = "";
            path = Globals.ThisAddIn.Application.ActivePresentation.Path;
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
        /*==========================================*/
        /*==========================================*/
        /*==========================================*/
        public string embedXLSX(string path, string title, string url, string user, string pwd)
        {
            string message="";
            string tmpdoc = "";
            object missing = System.Type.Missing;
            bool proceed = false;
            int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;

                         //title=title.Replace("/","");
                        // MessageBox.Show("title" + title);

                         try
                         {
                             tmpdoc = path + title;
                             downloadFile(url, tmpdoc, user, pwd);
                             proceed = true;

                         }
                         catch (Exception e)
                         {
                             MessageBox.Show("ERROR: "+e.Message);
                             string errorMsg = e.Message;
                             message = "error: " + errorMsg;
                         }

                         try
                         {
                             if (proceed)
                             {
                                 Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes.AddOLEObject(21, 105, 250, 250, "", tmpdoc, Microsoft.Office.Core.MsoTriState.msoFalse, "", 0, "", Microsoft.Office.Core.MsoTriState.msoFalse);
                             }
                         }
                         catch (Exception e)
                         {
                             MessageBox.Show("Error" + e.Message);
                             string errorMsg = e.Message;
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

        public string copyPasteSlideToActive(string tmpPath, string filename, string slideidx,string url,string user, string pwd, string retain)
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
                //MessageBox.Show("issue with download"+e.Message+e.StackTrace);
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            
            try
            {
                if (proceed)
                {
                    PPT.Presentation sourcePres = Globals.ThisAddIn.Application.Presentations.Open(sourcefile, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);
                    int num = Convert.ToInt32(slideidx);
                    copyPasteSlideToActiveSupport(sourcePres, num, retainformat);
                    sourcePres.Close();
                    sourcePres = null;
                }
            }
            catch(Exception e)
            {
                //MessageBox.Show("Unable to open: "+e.Message);
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                    
            }

            return message;
        }

        public string copyPasteSlideToActiveSupport(PPT.Presentation sourcePres, int slideidx, bool retain)
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
                        //MessageBox.Show("FAIL" + e.Message + "   " + e.StackTrace);
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
            imgDir = getTempPath() + fname;
            return imgDir;

        }

        public string useSaveFileDialog()
        {
            Prompt p = new Prompt();
            p.ShowDialog();
            string filename = p.pfilename;
            //MessageBox.Show(filename);
            if (!filename.EndsWith(".pptx"))
            {
                filename = filename + ".pptx";
                MessageBox.Show("filename");
            }
            return filename;
        }

        /*
        public string useSaveFileDialogOrig()
        {
            SaveFileDialog s = new SaveFileDialog();
           
            s.Filter = "PowerPoint Presentation (*.pptx)|*.pptx|All files (*.*)|*.*";
            s.DefaultExt = "pptx";
            s.AddExtension = true;
           
            s.ShowDialog();

            return s.FileName;
        }
        */

        private void downloadFile(string url, string sourcefile, string user, string pwd)
        {
            string message = "";
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

            //return message;
        }


        private void uploadData(string url, byte[] content, string user, string pwd)
        {
            string message = "";

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


           // return message;

        }

        private byte[] downloadData(string url, string user, string pwd)
        {
            //MessageBox.Show("downloading data");
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

        public string saveToML(string filename, string url, string user , string pwd)
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
                   // MessageBox.Show("message1 :" + message);
                }

                fs.Dispose();
                fs.Close();

            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                MessageBox.Show("message2 :" + message);
            }

            return message;
        }

        public string saveWithImages(string saveasdirectory, string saveasname, string url, string user, string pwd)
        {
            //dir parameter?  make optional in the javascript.  So you can save anywhere in ML.
            //remember to tied to filenames and mapping
           string message = "";

           try
           {

            PPT.Presentation pptx = Globals.ThisAddIn.Application.ActivePresentation;

            //string url = "http://localhost:8023/ppt/api/upload.xqy?uid=";

            string fullfilenamewithpath = "";
            string imgdirwithpath = "";
            string filename = "";

            fullfilenamewithpath = saveasdirectory + saveasname; // useSaveFileDialog()+".pptx";
            filename = fullfilenamewithpath.Split(new Char[] { '\\' }).Last();

            pptx.SaveAs(fullfilenamewithpath, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Microsoft.Office.Core.MsoTriState.msoFalse);
            //move url to ribbon as well

            //   url = url + "/" + filename;
            imgdirwithpath = convertFilenameToImageDir(fullfilenamewithpath);
            saveImages(imgdirwithpath, user, pwd);

            saveToML(fullfilenamewithpath, url, user, pwd);  //rename saveActivePresentation() - see excel


            /*
             * imgdirwithpath = convertFilenameToImageDir(fullfilenamewithpath);
            saveImages(imgdirwithpath, user, pwd);
             * */
          
           }
           catch (Exception e)
           {
               string errorMsg = e.Message;
               message = "error: " + errorMsg;
               //MessageBox.Show("Error" + e.Message);
           }

            return message;
        }

        /*
        public string saveWithImagesORIG(string saveasname, string saveas, string url)
        {
            //dir parameter?  make optional in the javascript.  So you can save anywhere in ML.
            //remember to tied to filenames and mapping

             string message = "";
             bool insuresaveas = false;
             if (saveas.Equals("true"))
                 insuresaveas = true;

            
            PPT.Presentation pptx = Globals.ThisAddIn.Application.ActivePresentation;

            //string url = "http://localhost:8023/ppt/api/upload.xqy?uid=";


            string path = pptx.Path;
            //dir parameter might be used here
            string filename = pptx.Name;
            string fullfilenamewithpath = "";
            string imgdirwithpath = "";
            string imgdir = "";

            if ((pptx.Name == null || pptx.Name.Equals("") || pptx.Path == null || pptx.Path.Equals(""))
                 ||insuresaveas)
            {
                fullfilenamewithpath = getTempPath() + saveasname + ".pptx"; // useSaveFileDialog()+".pptx";
                //MessageBox.Show("fullnamewithpath is now  " + fullfilenamewithpath);
                //here's where dir parameter might come in
                filename = fullfilenamewithpath.Split(new Char[] { '\\' }).Last();

                pptx.SaveAs(fullfilenamewithpath, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Microsoft.Office.Core.MsoTriState.msoFalse);
                url = url + "/" + filename;

                saveToML(fullfilenamewithpath, url);  //rename saveActivePresentation() - see excel


                imgdirwithpath = convertFilenameToImageDir(fullfilenamewithpath);
                //dir parameter?
                imgdir = imgdirwithpath.Split(new Char[] { '\\' }).Last();

                saveImages(imgdirwithpath);
               // pptx.SaveAs(imgdir, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsGIF, Microsoft.Office.Core.MsoTriState.msoFalse);

            }
            else
            {
               // MessageBox.Show("In the else");
                fullfilenamewithpath = path + "\\" + filename;
               // MessageBox.Show("Saving " + fullfilenamewithpath);
                pptx.Save();

                url = url + "/" + filename;
                try
                {
                    saveToML(fullfilenamewithpath, url);
                    //save to ML

                    imgdirwithpath = convertFilenameToImageDir(fullfilenamewithpath);
                    imgdir = imgdirwithpath.Split(new Char[] { '\\' }).Last();

                    saveImages(imgdirwithpath);
                }
                catch (Exception e)
                {
                    MessageBox.Show("Error" + e.Message);
                }
                
                //pptx.SaveAs(imgdir, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsGIF, Microsoft.Office.Core.MsoTriState.msoFalse);

            }


           // MessageBox.Show("fullnamewithpath:  "+fullfilenamewithpath + " imgdir: "+imgdirwithpath );

            return message;
        }
        */

        //need to make these changes
//public string saveImages(string imgdirwithpath, string mldir, string url, string user, string pwd)
        public string saveImages(string imgdirwithpath, string user, string pwd)
        {
            string message = "";
            string imgdir = imgdirwithpath.Split(new Char[] { '\\' }).Last();

            //name of folder with images, prepend with optional dir?
           // MessageBox.Show("IMGDIRWITHPATH.SPLIT.LAST: " + imgdir);
        
            MessageBox.Show("ImgDirWithPath: "+imgdirwithpath+"   ImageDir: " + imgdir);
            imgdir = "/" + imgdir; // +"/";
            PPT.Presentation ppt = Globals.ThisAddIn.Application.ActivePresentation;

            //need some try/catch action here ( and all over the place)
            if (Directory.Exists(imgdirwithpath))
            {
                string[] files = Directory.GetFiles(imgdirwithpath);
                foreach (string s in files)
                {
                    File.Delete(s);
                }
                Directory.Delete(imgdirwithpath);
            }

           // ppt.SaveAs(imgdirwithpath, PPT.PpSaveAsFileType.ppSaveAsGIF,Office.MsoTriState.msoFalse);
            ppt.SaveAs(imgdirwithpath, PPT.PpSaveAsFileType.ppSaveAsPNG,Office.MsoTriState.msoFalse);

            string[] imgfiles = Directory.GetFiles(imgdirwithpath);

            foreach (string i in imgfiles)
            {
               // MessageBox.Show("filename: " + i);
                string fname = i.Split(new Char[] { '\\' }).Last();
                string fileuri = imgdir + "/" + fname;
                //convert this uri to .pptx slide.xml
                MessageBox.Show("i"+i);
                MessageBox.Show("FileUri"+fileuri);
                //als get index from here
                // add as parameters for upload.xqy doc properties

              //  MessageBox.Show("fileuri to save :" + fileuri);
/*this is in pipeline now
                string parentprop = imgdir.Replace("_GIF", ".pptx");

                string slideprop = fname.Replace(".GIF", ".xml");
                slideprop = imgdir+"/ppt/slides/" + slideprop;
                slideprop = slideprop.Replace("_GIF", "_pptx_parts");
                slideprop = slideprop.Replace("Slide", "slide");
 * */

               // string slideprop = fileuri.Replace(".GIF", ".xml");
               // slideprop = slideprop.Replace("_GIF", "");

/* as is this
                string slideidx = fname.Replace("Slide", "");
                slideidx = slideidx.Replace(".GIF", "");
*/
              //  MessageBox.Show("properties: parent: " + parentprop + " slide: " + slideprop + " idx: " + slideidx);


                /*
                 * <pptx>/foo2.pptx</pptx>
                   <slide>/foo2_pptx_parts/ppt/slides/slide1.xml</slide>
                   <index>1</index>
                 * */
                //save to ml, pass imagesurl
                //link in .xqy to .pptx                                        //pptx name     //slide name          //slidename       
                //string url = "http://localhost:8023/ppt/api/upload.xqy?uid=" + fileuri+"&source="+parentprop+"&slide="+slideprop+"&idx="+slideidx;
                string url = "http://localhost:8023/ppt/api/upload.xqy?uid=" + fileuri;
                MessageBox.Show("url"+url);
                try
                {
                   
                    FileStream fs = new FileStream(i, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    int length = (int)fs.Length;
                    byte[] content = new byte[length];
                    fs.Read(content, 0, length);

                    try
                    {
                        uploadData(url, content,user,pwd);
                    }
                    catch (Exception e)
                    {
                        string errorMsg = e.Message;
                        message = "error: " + errorMsg;
                        MessageBox.Show("message1 :" + message);
                    }
                    
                    fs.Dispose();
                    fs.Close();
                   
                }
                catch (Exception e)
                {
                    string errorMsg = e.Message;
                    message = "error: " + errorMsg;
                    MessageBox.Show("message2 :" + message);
                }
            }

            //have the images, now have to get files and upload to ML
            //don't delete til we've copied to ML


            //Directory.Delete(imgdir);
            return message;
        }

        public string insertExcel()
        {
            string message = "";
            return message;
        }

        public string insertText(string txt)
        {
            //int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;

            //PPT.Shapes s = Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes;

            try
            {
                string orig =  Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange.Text;
                Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange.Text = orig + txt;
            }
            catch (Exception e)
            {
                MessageBox.Show("Please place select text insertion point with cursor.");
            }

           // PPT.TextRange tr = Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes[1].TextFrame.TextRange;
           // tr.Text = "FOOO";
            return "Foo";
           
        }

        public string insertTable() //parameterize rows, columns, vals
        {
           // MessageBox.Show("In addin");
            object missing = System.Type.Missing;
            int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
            PPT.Shape s = Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes.AddTable(2, 3,50,50,450,70);
            try
            {
                Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes.AddOLEObject(21, 105, 250, 250, "", @"C:\Workflow101.docx", Microsoft.Office.Core.MsoTriState.msoFalse, "", 0, "", Microsoft.Office.Core.MsoTriState.msoFalse);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error" + e.Message);
            }
            
            

            PPT.Table tbl = s.Table;
          
           // MessageBox.Show(tbl.Rows.Count + "here");
            PPT.Cell cell = tbl.Rows[1].Cells[1];
            cell.Shape.TextFrame.TextRange.Text = "Foo";
           // PPT.Shapes s = Globals.ThisAddIn.Application.ActivePresentation.Slides[sid].Shapes;
         

            return "foo";
        }

    }
}
