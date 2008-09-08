using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word=Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Security.Permissions;
//using System.Xml;
using System.IO;
//using DocumentFormat.OpenXml.Packaging; //OpenXML sdk
using Office = Microsoft.Office.Core;
using Microsoft.Win32;

namespace MarkLogic_WordAddin
{   
    [ComVisible(true)]
   // [ClassInterfaceAttribute(ClassInterfaceType.AutoDispatch)]
  //  [DockingAttribute(DockingBehavior.AutoDock)]
  //  [PermissionSetAttribute(SecurityAction.InheritanceDemand, Name = "FullTrust")]
  //  [PermissionSetAttribute(SecurityAction.LinkDemand, Name = "FullTrust")]

    public partial class UserControl1 : UserControl
    {
      //  private string tmpPath = "";                    
     //   private string propsFile = "";
        private string webUrl = "";
        private bool debug = false;
        private bool debugMsg = false;
        private string color = "";
        private string addinVersion = "@MAJOR_VERSION.@MINOR_VERSION@PATCH_VERSION";


        public UserControl1()
        {
            InitializeComponent();
   // tmpPath=Path.GetTempPath();
  //       propsFile=tmpPath+"OfficeProperties.txt";
 //CHANGED
 // bool configFileExists = checkForConnectionPropertiesFile();
            bool regEntryExists = checkUrlInRegistry();
            //MessageBox.Show("tmp path is" + tmpPath);
            //MessageBox.Show("propsFile is"+propsFile);
            //MessageBox.Show("file exists"+configFileExists);
            if (!regEntryExists)
            {
                //MessageBox.Show("Unable to find configuration info. Please insure OfficeProperties.txt exists in your system temp directory.  If problems persist, please contact your system administrator.");
                MessageBox.Show("                                   Unable to find configuration info. \n\r "+
                                " Please see the README for how to add configuration info for your system. \n\r "+
                                "           If problems persist, please contact your system administrator.");
            }
            else
            {
                //CHANGED
               // getConfigurationValues();
                color = TryGetColorScheme().ToString();
                webBrowser1.AllowWebBrowserDrop = false;
                webBrowser1.IsWebBrowserContextMenuEnabled = false;
                webBrowser1.WebBrowserShortcutsEnabled = false;
                webBrowser1.ObjectForScripting = this;
                webBrowser1.Navigate(webUrl);
                webBrowser1.ScriptErrorsSuppressed = true;
                

            }


        }

      //  checks for propsFile  tmpPath/OfficeProperties.txt
      //  private bool checkForConnectionPropertiesFile()
      //  {
      //      return File.Exists(propsFile);
      //  }

        private bool checkUrlInRegistry()
        {
            RegistryKey regKey1 = Registry.CurrentUser;
            regKey1 = regKey1.OpenSubKey(@"MarkLogicAddinConfiguration\Word");
            bool keyExists = false;
            if (regKey1 == null)
            {
                if(debugMsg)
                   MessageBox.Show("KEY IS NULL");

            }
            else
            {
                if(debugMsg)
                    MessageBox.Show("KEY IS: "+regKey1.GetValue("URL"));

                webUrl = (string)regKey1.GetValue("URL");
                if(!((webUrl.Equals(""))||(webUrl==null)))
                        keyExists = true;
            }
            return keyExists;
        }

        //read config info
   /*     private void getConfigurationValues()
        {
            TextReader tr = new StreamReader(propsFile);
            webUrl = tr.ReadLine();
            webUrl = webUrl.Trim();
            //string pwd = Encoding.Unicode.GetString(Convert.FromBase64String(tr.ReadLine()));
            //connPwd.Password = pwd;
            //connHost.Text = tr.ReadLine();
            //connPort.Text = tr.ReadLine();
            tr.Close();
            tr.Dispose();
    */
          /*  if (File.Exists(versionFile))
            {
                TextReader trv = new StreamReader(versionFile);
                connectLabel.Content = "Version: " + trv.ReadLine();
                trv.Close();
            }
          */

     //   }

        
        //used by CTPManager
        public Word.Document Document { get; set; }

        //used by CTPManager
        internal void Clear()
        {
        }

        //configuration info
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

        //start methods used by MarkLogicWordAddin.js

     /*   public String getConfiguration()
        {
            return addinVersion+"U+016000"+webUrl + "U+016000" + color ;
        }
     */

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

        public String getCustomPieceIds()
        {
            
            string ids = "";

            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                int count = doc.CustomXMLParts.Count;

                //ADDED THIS

                foreach (Office.CustomXMLPart c in doc.CustomXMLParts)
                {
                    if (c.BuiltIn.Equals(false))
                    {
                        ids += c.Id + "U+016000";

             
                    }
                }
                
                char[] tengwar = { 'U','+','0','1','6','0','0','0'};
                ids = ids.TrimEnd(tengwar);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                ids = "error";
            }

            if (debug)
                ids = "error";

            //MessageBox.Show(ids);
            return ids;
        }


        public String getCustomPiece(string id)
        {

            string custompiecexml = "";

            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Office.CustomXMLPart cx = doc.CustomXMLParts.SelectByID(id);

                if (cx != null)
                    custompiecexml = cx.XML;

                /*another way (used until I discovered SelectByID(id) above)
                  foreach (Office.CustomXMLPart c in doc.CustomXMLParts)
                  {
                      if (c.BuiltIn.Equals(false) && c.Id.Equals(id))
                      {
                          Office.CustomXMLNode x = c.DocumentElement;
                          custompiecexml = x.XML;
                      }
                
                  }
                 */
            }catch(Exception e){
                string errorMsg = e.Message;
                custompiecexml = "error";
            }

            if (debug)
                custompiecexml = "error";

            return custompiecexml;

        }

        public String addCustomPiece(string custompiecexml)
        {
            string newid = "";
            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Office.CustomXMLPart cx = doc.CustomXMLParts.Add(String.Empty, new Office.CustomXMLSchemaCollectionClass());
                cx.LoadXML(custompiecexml);
                newid = cx.Id;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                //should we display error message here? leaving in js for now.
                newid = "error";
            }
            if (debug)
                newid = "error";
            return newid;

        }

        public String deleteCustomPiece(string id)
        {
            string message = "";
            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                foreach (Office.CustomXMLPart c in doc.CustomXMLParts)
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

        //currently no way to replace without delete,add, get new id

        public String getSelection()
        {
            string wpml = "";
            try
            {
                Word.Range rng = Globals.ThisAddIn.Application.Selection.Range;
                int stTst = rng.Start;
                int edTst = rng.End;
                string xmlizable = "";
               

                if (stTst < edTst)
                {
                    rng.Select();
                    xmlizable = Globals.ThisAddIn.Application.Selection.WordOpenXML; // wordApp.Selection.WordOpenXML;  //instead of .Text
                    wpml = Transform.ConvertToWPMLDelimitedFromText(xmlizable);
                }
                else
                {
                    wpml = "";
                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                wpml = "error";
            }
            
            if(debugMsg)
               MessageBox.Show("returning wpml: " + wpml);

            if (debug)
                wpml = "error";

            return wpml;

        }

        //returns the entire styles.xml from the active package
        public String getActiveDocStylesXml()
        {
            string stylesxml="";
            try
            {
                string docxml = Globals.ThisAddIn.Application.ActiveDocument.WordOpenXML;
                stylesxml = Transform.GetStylesXmlFromCurrentDoc(docxml);

               if (debugMsg)
               {
                    TextWriter tw = new StreamWriter(@"C:\styles.xml");
                    tw.WriteLine(stylesxml);
                    tw.Close();
               }


            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                stylesxml = "error";
            }

            if (debug)
                stylesxml = "error";

            return stylesxml;

        }

        //returns the style for the current block
        public String getRangePreview()
        {
            string wpml = "";
            try
            {

                Word.Range rng = Globals.ThisAddIn.Application.Selection.Range;
                int stTst = rng.Start;
                int edTst = rng.End;
                string xmlizable = "";

                if (stTst < edTst)
                {

                    rng.Select();
                    xmlizable = Globals.ThisAddIn.Application.Selection.WordOpenXML; // wordApp.Selection.WordOpenXML;  //instead of .Text
                    wpml = Transform.ConvertToWPMLFromTextFinalNode(xmlizable);
                }
                else
                {
                    int origStart = rng.Start;
                    int origEnd = rng.End;
                    int newStart = origEnd;// -1;
                    int newEnd = origEnd;
                    object startLocation = newStart;
                    object endLocation = newEnd;

                    //need to grab range where cursor is for property preview
                    rng = Globals.ThisAddIn.Application.Selection.Range.Sentences[1];
                    rng.Select();
                    xmlizable = Globals.ThisAddIn.Application.Selection.WordOpenXML;
                    wpml = Transform.ConvertToWPMLFromText(xmlizable);

                    rng = Globals.ThisAddIn.Application.ActiveDocument.Range(ref startLocation, ref endLocation);
                    rng.Select();

                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                wpml = "error";
            }

            if(debugMsg)
               MessageBox.Show("returning wpml: " + wpml);

            if(debug)
               wpml = "error";

            return wpml;
        }

        //have to remove namespace to be able to insert
        public String removeNamespaces(string xml)
        {   
            string tmp = "";
            tmp = xml.Replace(" xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"", "");
            tmp = tmp.Replace(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"", "");
            tmp = tmp.Replace(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"", "");
            tmp = tmp.Replace(" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"", "");
            tmp = tmp.Replace(" xmlns:v=\"urn:schemas-microsoft-com:vml\"", "");
            tmp = tmp.Replace(" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"", "");
            tmp = tmp.Replace(" xmlns:w10=\"urn:schemas-microsoft-com:office:word\"", "");
            tmp = tmp.Replace(" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", "");
            tmp = tmp.Replace(" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\"", "");
          //  tmp = tmp.Replace("   ", "");

            return tmp;

        }

        public String insertBlockContent(string blockContent, string stylesXml)
        {
            string clean = "";
            //MessageBox.Show("IN insertBlockContent");
            clean = removeNamespaces(blockContent);
            string msg = insertBlock(blockContent, stylesXml);

            return msg;

        }

        public String insertBlock(String blockContent, String stylesXml)
        {
            string message = "";
            string wpml = blockContent;
            string newStyle = stylesXml;
            object missing = System.Reflection.Missing.Value;

            //Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            //check to see if range selected
            Word.Range testrng = Globals.ThisAddIn.Application.Selection.Range;
            int selectedRangeStart = testrng.Start;
            int selectedRangeEnd = testrng.End;


            //get index of paragraph from cursor to pass for XPath
            //ActiveDocument.Range(0, Selection.Paragraphs(1).Range.End).Paragraphs.Count
            object start = 0;
            object paraend = Globals.ThisAddIn.Application.Selection.Paragraphs[1].Range.End;
            int paraidx = Globals.ThisAddIn.Application.ActiveDocument.Range(ref start, ref paraend).Paragraphs.Count;

            object sentend = Globals.ThisAddIn.Application.Selection.Sentences[1].End;
            int sentidx = Globals.ThisAddIn.Application.ActiveDocument.Range(ref start, ref sentend).Sentences.Count;

            int currentcharidx = Globals.ThisAddIn.Application.Selection.End;


            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            try
            {
                if (!(selectedRangeStart < selectedRangeEnd))
                {
                    doc.Application.Selection.InsertXML(Transform.ConvertToWPMLBlock(newStyle, wpml, paraidx, sentidx, currentcharidx), ref missing); 
                }
                else
                {
                    doc.Application.Selection.InsertXML(Transform.ConvertToWPMLBlock(newStyle, wpml, selectedRangeStart, selectedRangeEnd), ref missing); 
                }
            }
            catch (Exception e)
            {
                string errMsg = e.Message;
                //System.Windows.Forms.MessageBox.Show("Error in InsertBlock.\r\n\r\nUnable To Insert At This Location.\r\n\r\n" + ef.Message + ef.StackTrace);
                message = "error";
            }

            if (debug)
                message = "error";

            return message;
        }

   /*     public static void AddImagePart(string document, string fileName)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                using (FileStream stream = new FileStream(fileName, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }
            }
        }
    * */

    }
}
