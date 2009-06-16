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

     /*   private bool checkUrlInRegistry()
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
      * */
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
                ids = "error";
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
                custompiecexml = "error";
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
                newid = "error";
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
                message = "error";
            }

            if (debug)
                message = "error";

            return message;

        }

        public Image byteArrayToImage(byte[] byteArrayIn)
        {
            MemoryStream ms = new MemoryStream(byteArrayIn);
            Image returnImage = Image.FromStream(ms);
            return returnImage;
        }

        public byte[] imageToByteArray(System.Drawing.Image imageIn)
        {
            MemoryStream ms = new MemoryStream();
            imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Gif);
            return ms.ToArray();
        }

        public String insertImage(string imageuri, string uname, string pwd)
        {
            object missing = Type.Missing;
            //MessageBox.Show("Adding Image");
            string message = "";

            System.Net.WebClient Client = new System.Net.WebClient();
            Client.Credentials = new System.Net.NetworkCredential(uname, pwd);
            byte[] bytearray = Client.DownloadData(imageuri);
            Image img = byteArrayToImage(bytearray);
            //Image img = Image.FromFile(@"C:\gijoe_destro.jpg");


            PPT.Slide slide = (PPT.Slide)Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            Clipboard.SetImage(img);
            slide.Shapes.Paste();
            Clipboard.Clear();
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

        public String openPPTX(string path, string title, string url, string user, string pwd)
        {
            // MessageBox.Show("in the addin path:"+path+  "      title"+title+ "   uri: "+url+"user"+user+"pwd"+pwd);
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

                //Excel.Workbook wb = Globals.ThisAddIn.Application.Workbooks.Open(tmpdoc, missing, false, missing, missing, missing, true, missing, missing, true, true, missing, missing, missing, missing);

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
                string origmsg = "A document with the name '" + title + "' is already open. You cannot open two documents with the same name, even if the documents are in different \nfolders. To open the second document, either close the document that's currently open, or rename one of the documents.";
                MessageBox.Show(origmsg);
                string errorMsg = e.Message;
                message = "error: " + errorMsg;

            }

            return message;
        }

        public string copyPasteSlideToActive(string tmpPath,string filename, string slideidx,string url,string user, string pwd)
        {
            //MessageBox.Show("In function tmppath"+tmpPath+" filename: "+filename+" slideidx"+slideidx+ "url: "+url+"user: "+user+"pwd"+pwd);

            string message = "";
            object missing = Type.Missing;
            string sourcefile = "";
            string path = getTempPath();

            try
            {
                System.Net.WebClient Client = new System.Net.WebClient();
                Client.Credentials = new System.Net.NetworkCredential(user, pwd);
              
                sourcefile = path + filename;
                Client.DownloadFile(url, sourcefile);
            }
            catch (Exception e)
            {
                MessageBox.Show("issue with download");
            }

            try
            {
                PPT.Presentation sourcePres = Globals.ThisAddIn.Application.Presentations.Open(sourcefile, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);
                int num = Convert.ToInt32(slideidx);
                copyPasteSlideToActiveSupport(sourcePres,num );
                sourcePres.Close();
                sourcePres = null;
            }
            catch
            {
                MessageBox.Show("Unable to open");
                    
            }

            return message;
        }
        public string copyPasteSlideToActiveSupport(PPT.Presentation sourcePres, int slideidx)
        {
            //arguments need to include slide(s) to be inserted ..
            // user, pwd, url, title, tmpath(?), retainsourceformatting 

            //get index of starter slide and reset at end of function?
            //don't have to worry about if just inserting one slide at a time.

            MessageBox.Show("Copy Pasting files  --");
            //MessageBox.Show("1: "+GC.MaxGeneration);

            //try getting this from server
            // string sourcefile = @"C:\Aven_MarkLogicUserConference2009Exceling.pptx";

            PPT.Presentation activePres = Globals.ThisAddIn.Application.ActivePresentation;
            //MessageBox.Show("3: " + GC.MaxGeneration + "activepresegen: " + GC.GetGeneration(activePres));
            // PPT.Presentation sourcePres = Globals.ThisAddIn.Application.Presentations.Open(sourcefile, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);
            //activePres.SlideMaster.BackgroundStyle = sourcePres.SlideMaster.BackgroundStyle;

            PPT.Slides activeSlides = activePres.Slides;
            PPT.Slides sourceSlides = sourcePres.Slides;

            for (int x = 1; x < sourceSlides.Count; x++)
            {
                int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
                int id = sourceSlides[x].SlideID;

                if (sourceSlides[x].SlideIndex == slideidx)
                {
                    //MessageBox.Show(id+"");
                    sourceSlides.FindBySlideID(id).Copy();
                    //sourcePres.SlideMaster.Background.
                    //activePres.Application.ActiveWindow.View.PasteSpecial();
                    //activeSlides.Paste(x);
                    try
                    {
                        //int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
                        // MessageBox.Show("Idx before:  " + Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex);

                        activeSlides.Paste(sid).FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
                        //if need to pull in master, then (also don't set follow master background above

                        Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[sid].Select();
                        PPT.SlideRange sr = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange;
                        sr.Design = sourcePres.SlideMaster.Design;
                        Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[sid + 1].Select();
                        ///sr.BackgroundStyle = sourceSlides.FindBySlideID(id).BackgroundStyle;//sourcePres.SlideMaster.Background;
                        //  sr.ColorScheme = sourceSlides.FindBySlideID(id).ColorScheme;//sourcePres.SlideMaster.ColorScheme;
                        // sr.DisplayMasterShapes = //Microsoft.Office.Core.MsoTriState.msoTrue;

                        //activeSlides[x].Background.BackgroundStyle = sourceSlides.FindBySlideID(id).Background.BackgroundStyle;
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("FAIL" + e.Message + "   " + e.StackTrace);
                    }
                }


            }



            MessageBox.Show("returning foo");
            return "foo";

        }

        public string convertFilenameToImageDir(string filename)
        {
            string imgDir = "";
            string tmpDir = "";
            string fname = "";

            string[] split = filename.Split(new Char[] { '\\' });
            fname = split.Last();
            tmpDir = filename.Replace(fname, "");
            fname = fname.Replace(".pptx", "_pptx_parts_GIF");

            //imgDir = tmpDir + fname;
            imgDir = getTempPath() + fname;
            MessageBox.Show("imgdir: "+imgDir);
            return imgDir;

        }

        public string useSaveFileDialog()
        {

            SaveFileDialog s = new SaveFileDialog();
            s.Filter = "PowerPoint Presentation (*.pptx)|*.pptx|All files (*.*)|*.*";
            s.DefaultExt = "pptx";
            s.AddExtension = true;
            s.ShowDialog();

            return s.FileName;
        }

        public string uploadFile(string url, byte[] content)
        {

            System.Net.WebClient Client = new System.Net.WebClient();
            Client.Headers.Add("enctype", "multipart/form-data");
            Client.Headers.Add("Content-Type", "application/octet-stream");
            Client.Credentials = new System.Net.NetworkCredential("oslo", "oslo");

            Client.UploadData(url, "POST", content);
            Client.Dispose();

            return "foo";

        }

        public string saveToML(string filename, string url)
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
                    uploadFile(url, content);
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

            return message;
        }

        public string saveWithImages()
        {
            //dir parameter?  make optional in the javascript.  So you can save anywhere in ML.
            //remember to tied to filenames and mapping
           
            string message = "";
            PPT.Presentation pptx = Globals.ThisAddIn.Application.ActivePresentation;

            string url = "http://localhost:8023/ppt/api/upload.xqy?uid=";


            string path = pptx.Path;
            //dir parameter might be used here
            string filename = pptx.Name;
            string fullfilenamewithpath = "";
            string imgdirwithpath = "";
            string imgdir = "";

            if (pptx.Name == null || pptx.Name.Equals("") || pptx.Path == null || pptx.Path.Equals(""))
            {
                fullfilenamewithpath = useSaveFileDialog();
                //here's where dir parameter might come in
                filename = fullfilenamewithpath.Split(new Char[] { '\\' }).Last();

                pptx.SaveAs(fullfilenamewithpath, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Microsoft.Office.Core.MsoTriState.msoFalse);
                url = url + "/" + filename;

                saveToML(fullfilenamewithpath, url);


                imgdirwithpath = convertFilenameToImageDir(fullfilenamewithpath);
                //dir parameter?
                imgdir = imgdirwithpath.Split(new Char[] { '\\' }).Last();

                saveImages(imgdirwithpath);
               // pptx.SaveAs(imgdir, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsGIF, Microsoft.Office.Core.MsoTriState.msoFalse);

            }
            else
            {
                fullfilenamewithpath = path + "\\" + filename;
                pptx.Save();
                url = url + "/" + filename;
                saveToML(fullfilenamewithpath, url);
                //save to ML

                imgdirwithpath = convertFilenameToImageDir(fullfilenamewithpath);
                imgdir = imgdirwithpath.Split(new Char[] { '\\' }).Last();

                saveImages(imgdirwithpath);
                
                //pptx.SaveAs(imgdir, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsGIF, Microsoft.Office.Core.MsoTriState.msoFalse);

            }


            MessageBox.Show("fullnamewithpath:  "+fullfilenamewithpath + " imgdir: "+imgdirwithpath );

            return message;
        }

        public string saveImages(string imgdirwithpath)
        {
            string message = "";
            string imgdir = imgdirwithpath.Split(new Char[] { '\\' }).Last();
            MessageBox.Show("IMGDIRWITHPATH.SPLIT.LAST: " + imgdir);
            imgdir = "/" + imgdir+"/";
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

            ppt.SaveAs(imgdirwithpath, PPT.PpSaveAsFileType.ppSaveAsGIF,Office.MsoTriState.msoFalse);

            string[] imgfiles = Directory.GetFiles(imgdirwithpath);

            foreach (string i in imgfiles)
            {
                //MessageBox.Show("filename: " + i);
                string fname = i.Split(new Char[] { '\\' }).Last();
                string fileuri = imgdir + fname;
                MessageBox.Show("fileuri to save :" + fileuri);

                //save to ml, pass imagesurl
                string url = "http://localhost:8023/ppt/api/upload.xqy?uid=" + fileuri;

                try
                {
                   
                    FileStream fs = new FileStream(i, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    int length = (int)fs.Length;
                    byte[] content = new byte[length];
                    fs.Read(content, 0, length);

                    try
                    {
                        uploadFile(url, content);
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
//====================================================================================================
//====================================================================================================
//====================================================================================================
    //  public static bool CopySlidesFromPPT(string sourcefile, string dstfile, out string exmsg)
        public string copyPasteSlideToActiveSupportBACKUP(PPT.Presentation sourcePres)
        {
            MessageBox.Show("Copy Pasting files  --");
            //MessageBox.Show("1: "+GC.MaxGeneration);

            //try getting this from server
// string sourcefile = @"C:\Aven_MarkLogicUserConference2009Exceling.pptx";

             PPT.Presentation activePres = Globals.ThisAddIn.Application.ActivePresentation;
            //MessageBox.Show("3: " + GC.MaxGeneration + "activepresegen: " + GC.GetGeneration(activePres));
// PPT.Presentation sourcePres = Globals.ThisAddIn.Application.Presentations.Open(sourcefile, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);

            //activePres.SlideMaster.BackgroundStyle = sourcePres.SlideMaster.BackgroundStyle;
          
            PPT.Slides activeSlides = activePres.Slides;
            PPT.Slides sourceSlides = sourcePres.Slides;

            for (int x = 1; x < sourceSlides.Count; x++)
            {
                int id = sourceSlides[x].SlideID;
                //MessageBox.Show(id+"");
                sourceSlides.FindBySlideID(id).Copy();
                //sourcePres.SlideMaster.Background.
                //activePres.Application.ActiveWindow.View.PasteSpecial();
                //activeSlides.Paste(x);
                try
                {
                    int sid = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex;
                   // MessageBox.Show("Idx before:  " + Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.SlideIndex);

                    activeSlides.Paste(sid).FollowMasterBackground = Microsoft.Office.Core.MsoTriState.msoFalse;
          //if need to pull in master, then (also don't set follow master background above
         
             Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[sid].Select();
             PPT.SlideRange sr = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange;
             sr.Design = sourcePres.SlideMaster.Design;
             Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides[sid+1].Select();
                    ///sr.BackgroundStyle = sourceSlides.FindBySlideID(id).BackgroundStyle;//sourcePres.SlideMaster.Background;
                  //  sr.ColorScheme = sourceSlides.FindBySlideID(id).ColorScheme;//sourcePres.SlideMaster.ColorScheme;
                   // sr.DisplayMasterShapes = //Microsoft.Office.Core.MsoTriState.msoTrue;

                 //activeSlides[x].Background.BackgroundStyle = sourceSlides.FindBySlideID(id).Background.BackgroundStyle;
                }
                catch (Exception e)
                {
                    MessageBox.Show("FAIL"+e.Message+"   "+e.StackTrace);
                }
                

            }

  

            MessageBox.Show("returning foo");
            return "foo";

        }

        //missing template style BLERG!
        public string copySlideToActive()
        {
            MessageBox.Show("Saving files 1");
            string sourcefile = @"C:\Aven_MarkLogicUserConference2009Exceling.pptx";
            //PPT.Presentation p = Globals.ThisAddIn.Application.ActivePresentation;

            PPT.Application ppa = new PPT.ApplicationClass();
            PPT.Presentations ppp = ppa.Presentations;
            //PPT.Presentation ppmp = null;

            PPT.Presentation ppmp = Globals.ThisAddIn.Application.ActivePresentation;

            PPT.Slides ppms = ppmp.Slides;
            
            

            PPT.Presentation ppps = ppp.Open( sourcefile, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);
            
            ppms.InsertFromFile( sourcefile, ppms.Count, 1, ppps.Slides.Count);
           

            //ppmp.SlideMaster.CustomLayouts.Add( ppps.SlideMaster.CustomLayouts);
            //ppmp.SlideMaster.
            ////ppmp.SlideMaster.Background.BackgroundStyle = ppps.SlideMaster.Background.BackgroundStyle;
           // ppmp.SlideMaster.ColorScheme = ppps.SlideMaster.ColorScheme;
           // ppmp.HandoutMaster.BackgroundStyle = ppps.HandoutMaster.BackgroundStyle;
            try
            {
                ppmp.SlideMaster.BackgroundStyle = ppps.SlideMaster.BackgroundStyle;
            }
            catch (Exception e)
            {
                MessageBox.Show("FAIL");
            }
            finally
            {

                ppps.Close();
            }
            

                                //ppmp.Close();
                    //
                    // Release the COM object holding the merged presentation
                    //
                    //Marshal.ReleaseComObject(ppmp);
                   // ppmp = null;
                    //
                    // Release the COM object holding the presentations
                    //
            Marshal.ReleaseComObject(ppp);
            ppp = null;
                    //
                    // Release the COM object holding the powerpoint application
                    //
            Marshal.ReleaseComObject(ppa);
            ppa = null;
                

            return "foo";

            
        }


        public string CopySlidesFromPPT()
        {
            MessageBox.Show("Saving files 2");
            string sourcefile = @"C:\Aven_MarkLogicUserConference2009Exceling.pptx";
            string dstfile = @"C:\JetBlue case study r6.pptx";
            string exmsg="";
            bool success = false;

            //
            // Initialise the exception message
            //
            exmsg = "";
            //
            // Create a link to the PowerPoint object model
            //
            PPT.Application ppa = new PPT.ApplicationClass();
            PPT.Presentations ppp = ppa.Presentations;
            PPT.Presentation ppmp = null;
            //
            // If the destination presentation exists on disk, load it so
            // that we can append the new slides
            //
            if (File.Exists(dstfile) == true)
            {
                try
                {
                    //
                    // Try and open the destination presentation
                    //
                    ppmp = ppp.Open(dstfile, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse);
                }
                catch (Exception ex)
                {
                    ppmp = null;
                    exmsg = ex.Message;
                }
            }
            else
            {
                //
                // Create a new presentation
                //
                try
                {
                    ppmp = ppp.Add(Microsoft.Office.Core.MsoTriState.msoFalse);
                }
                catch (Exception ex)
                {
                    ppmp = null;
                    exmsg = ex.Message;
                }
            }
            //
            // Do we have a master presentation ?
            //
            if (ppmp != null)
            {
                //
                // Point to the slides in the master presentation
                //
               PPT.Slides ppms = ppmp.Slides;
                try
                {
                    try
                    {
                        //
                        // Open the source presentation
                        //
                        PPT.Presentation ppps = ppp.Open(sourcefile, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);
                        try
                        {
                            //
                            // Insert the source slides onto the end of the merge presentation
                            //
                            ppms.InsertFromFile(sourcefile, ppms.Count, 1, ppps.Slides.Count);
                            //
                            // Save the merged presentation back to disk
                            //
                            ppmp.SaveAs(dstfile, PPT.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoFalse);
                            //
                            // Signal success
                            success = true;
                        }
                        finally
                        {
                            //
                            // Close the source presentation
                            //
                            ppps.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        exmsg = ex.Message;
                    }
                }
                finally
                {
                    //
                    // Ensure the merge presentation is closed
                    //
                    ppmp.Close();
                    //
                    // Release the COM object holding the merged presentation
                    //
                    Marshal.ReleaseComObject(ppmp);
                    ppmp = null;
                    //
                    // Release the COM object holding the presentations
                    //
                    Marshal.ReleaseComObject(ppp);
                    ppp = null;
                    //
                    // Release the COM object holding the powerpoint application
                    //
                    Marshal.ReleaseComObject(ppa);
                    ppa = null;
                    ppa.Quit();
                }
            }
            return "TEST";
        }

        /*
        public String insertImage(string imageuri, string imagename)
        {
            object missing = Type.Missing;
            MessageBox.Show("Adding Image");
            string message = "";

            System.Net.WebClient Client = new System.Net.WebClient();
            Client.Credentials = new System.Net.NetworkCredential("zeke", "zeke");
            byte[] bytearray = Client.DownloadData(imageuri);
            Image img = byteArrayToImage(bytearray);
            //Image img = Image.FromFile(@"C:\gijoe_destro.jpg");


            PPT.Slide slide = (PPT.Slide)Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

            Clipboard.SetImage(img);
            slide.Shapes.Paste();
            Clipboard.Clear();
            return message;
        }
         * */

        public String addSlide()
        {
      
            MessageBox.Show("IN ADDIN");
            string message = "foo";
            object missing = Type.Missing;
            //string message="";
            string filename = @"C:\MarkLogic Connector for SharePoint r1,1.pptx";
         //   PPT.Application ppa = new PPT.ApplicationClass();
         //   ppa.Presentations.Open(filename,Microsoft.Office.Core.MsoTriState.msoFalse,Microsoft.Office.Core.MsoTriState.msoFalse,Microsoft.Office.Core.MsoTriState.msoFalse);
         //   ppa.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
         //   ppa.Activate();

            try
            { 
              
                PPT.Presentation actP = Globals.ThisAddIn.Application.ActivePresentation;
                

                PPT.Application ppa = new PPT.ApplicationClass();
              
                ppa.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                ppa.Presentations.Open(filename, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue);

                PPT.Presentation sourceP = ppa.ActivePresentation;
                



                PPT.Slide slide = actP.Slides[1];
                //PPT.Slide slide = actP.Slides.Add(actP.Slides.Count, Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);
              //  PPT.Slide s2 = actP.Slides[slide.SlideIndex];
                
                slide.Application.Activate();
                /*
                 * objSourcePresentation.Slides(SlideID).Copy()
                   objDestinationPresentation.Slides.Paste
                 * */
                MessageBox.Show(sourceP.Slides.Count+"");
                sourceP.Slides[1].Copy();
                //slide.BackgroundStyle = sourceP.SlideMaster.BackgroundStyle;

                actP.Slides.Paste(slide.SlideIndex);
                PPT.Slide s2 = actP.Slides[slide.SlideIndex - 1];
                s2.CustomLayout = sourceP.Slides[1].CustomLayout;

               
                
             //   actP.Application.Activate();
                //sourceP.Close();
                //slide.CustomLayout = sourceP.Slides[1].CustomLayout;
         //       actP.Slides.Paste(slide.SlideIndex);
                //slide.BackgroundStyle = sourceP.SlideMaster.BackgroundStyle;

                
                //actP.SlideMaster.BackgroundStyle = sourceP.SlideMaster.BackgroundStyle;
                //sourceP.Close();


                //slide = ppa.Presentations[1].Slides[4];
                //PPT.Application ppa = Globals.ThisAddIn.Application;
                //PPT.Presentations ppts = Globals.ThisAddIn.Application.Presentations;
                //ppts.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
                //ppts.Open(filename, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue);
                
               

               // PPT.Application ppa = new PPT.ApplicationClass();
               // ppa.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
               // PPT.Presentation ppts = Globals.ThisAddIn.Application.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
              
               //PPT.Presentations ppp = ppa.Presentations;
                //PPT.Presentation ppmp = null;
                //ppa.Presentations.Open(filename, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue);
               // ppts.Open(filename, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue);
               // PPT.Presentation p1 = Globals.ThisAddIn.Application.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
               // p1 = ppa.Presentations[1];
               // p1.Application.Activate();
                //ppa.Activate();
            }
            catch (Exception e)
            {
                MessageBox.Show("ERROR" + e.Message + "==================" + e.StackTrace);
            }
            

//PPTApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue; ;

          //  PPT.Presentations ppp = ppa.Presentations;
          //  ppp.Open(@"http://localhost:8011/MarkLogic Connector for SharePoint r1,1.pptx"

            //PPT.Application objSourcePresentation;
            
           // = File.Open(@"http://localhost:8011/MarkLogic Connector for SharePoint r1,1.pptx", FileMode.Open);


  //          'copies the source slide to the clipboard
//objSourcePresentation.Slides(SlideID).Copy()


//'appends the slide from the clipboard to the end of the other presentation
//objDestinationPresentation.Slides.Paste


            return message;

        }

        //TODO:
        //pass filename, imagename, uri - want to use client download to tmp file
        //insert image
        //delete tmp file
 /*      public String insertImageORIG(string imageuri, string imagename)
        {
            object missing = Type.Missing;
            MessageBox.Show("Adding Image");
            string message = "";
          //PPT.Slide s = Globals.ThisAddIn.Application.ActivePresentation.Slides[Globals.ThisAddIn.Application.ActivePresentation.Slides.];
//ONE WAY
           // try this Current slide? gijoe too
 PPT.Slide slide = (PPT.Slide)Globals.ThisAddIn.Application.ActiveWindow.View.Slide;// app.ActiveWindow.View.Slide;

//ADDING SLIDE
          PPT.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
//ADDING SLIDE PPT.Slide slide =
//ADDING SLIDE presentation.Slides.Add(
//ADDING SLIDE presentation.Slides.Count + 1,
//ADDING SLIDE PPT.PpSlideLayout.ppLayoutPictureWithCaption);

            //can get this from byte array
          //Image img = Image.FromFile(@"C:\test.png");
          Image img = Image.FromFile(@"C:\gijoe_destro.jpg");
         
          Clipboard.SetImage(img);
          //slide.Shapes.Paste();
            //how to add to current slide?
          slide.Shapes.Paste();
         // presentation.Slides[1].Shapes.Paste();

          //  richTextBox1.SelectionStart = 0;
          //  richTextBox1.Paste();

          Clipboard.Clear();

//ONE WAY PPT.Shape shape = slide.Shapes[2];
// ONE WAY slide.Shapes.AddPicture(@"C:\test.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
//shape.Left, shape.Top, shape.Width, shape.Height);
//ONE WAY slide.Select();
            
           

                //shape.Left, shape.Top, shape.Width, shape.Height);
             

           // slide.Shapes.AddPicture();
            //PPT.Presentation s = Globals.ThisAddIn.Application.ActivePresentation;
            //s.Application.ActiveWindow.ActivePane;
           // PPT.Slide s = Globals.ThisAddIn.Application.ActivePresentation;
           // Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

   //         System.Net.WebClient Client = new System.Net.WebClient();
   //         Client.Credentials = new System.Net.NetworkCredential("zeke", "zeke");

   //         byte[] bytearray = Client.DownloadData(imageuri);
   //         Image img = byteArrayToImage(bytearray);


            //place on clipboard
   //         System.Windows.Forms.Clipboard.SetImage(img);
           
           // Globals.ThisAddIn.Application.Selection.Range.Paste();
 
            return message;
        }

  */



    }
}
