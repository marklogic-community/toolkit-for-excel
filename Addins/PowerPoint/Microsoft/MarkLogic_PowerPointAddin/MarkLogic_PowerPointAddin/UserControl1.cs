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
            MessageBox.Show("Adding Image");
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
        public String insertImageORIG(string imageuri, string imagename)
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

  



    }
}
