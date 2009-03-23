/*Copyright 2008 Mark Logic Corporation

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
 * UserControl1.cs - the api called from MarkLogicWordAddin.js.  The methods here map directly to functions in the .js.
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

using Word=Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.IO;
using Office = Microsoft.Office.Core;
using Microsoft.Win32;
using System.Windows.Forms.Integration;
using DocumentFormat.OpenXml.Packaging;


namespace MarkLogic_WordAddin
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
            webUrl = ac.getWebURL();

            if (webUrl.Equals(""))
            {
                MessageBox.Show("                                            Unable to find configuration info. \n\r "+
                                " Please see the installation instructions for how to add configuration info to your system. \n\r "+
                                "                   If problems persist, please contact your system administrator.");
            }
            else
            {

                color = TryGetColorScheme().ToString();
                webBrowser1.AllowWebBrowserDrop = false;
                webBrowser1.IsWebBrowserContextMenuEnabled = false;
                webBrowser1.WebBrowserShortcutsEnabled = true;
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
                 
                 {
                     webBrowser1.Parent.Focus();
                     webBrowser1.Document.Focus();
                  

                 }
                 
              }
        }

      //public Word.Document Document { get; set; }

      //internal void Clear()
      //{
      //}

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
            //assume default - theme registry key not always set on install of Office
            //set once user sets color scheme manually from button
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

        public String getBrowserURL()
        {
            return webUrl;
        }

        public String getCustomXMLPartIds()
        {
            
            string ids = "";

            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                int count = doc.CustomXMLParts.Count;

                foreach (Office.CustomXMLPart c in doc.CustomXMLParts)
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
                ids = "error: "+errorMsg;
            }

            if (debug)
                return "error: TESTING ERRORS";

            return ids;
        }


        public String getCustomXMLPart(string id)
        {

            string custompiecexml = "";

            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                Office.CustomXMLPart cx = doc.CustomXMLParts.SelectByID(id);

                if (cx != null)
                    custompiecexml = cx.XML;

                /*another way; how should we expose built-ins? do we want to?
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
                custompiecexml = "error: "+errorMsg;
            }

            if (debug)
                custompiecexml = "error: TESTING ERRORS";

            return custompiecexml;

        }

        public String addCustomXMLPart(string custompiecexml)
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
                newid = "error: "+errorMsg;
            }
            if (debug)
                newid = "error: testing Errors";

            return newid;

        }

        public String deleteCustomXMLPart(string id)
        {
            string message = "";
            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                foreach (Office.CustomXMLPart c in doc.CustomXMLParts)
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
                message = "error: "+errorMsg;
            }

            if (debug)
                message = "error: Testing errors.";

            return message;
             
        }

        public String getSelection(int idx)
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
                    wpml = Transform.ConvertToWPMLFromTextIdx(xmlizable, idx);
                }
                else
                {
                    wpml = "";
                }
            }


            catch (Exception e)
            {
                string errorMsg = e.Message;
                wpml = "error: " + errorMsg;
            }

            if (debugMsg)
                MessageBox.Show("returning wpml: " + wpml);

            if (debug)
                wpml = "error: Testing errors";

            return wpml;

        }

        //this works great if you're positive all references for the other pieces in
        //package are already resolved (templates).  Useful for enriching entire doc with
        //w:customXml, etc.  
        public String setActiveDocXml(string wpml)
        {
            string docxml = "";
            object missing = System.Reflection.Missing.Value;

            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                doc.Application.Selection.InsertXML(Transform.ConvertToWPML(wpml), ref missing);
            } 
            catch (Exception e)
            {
                string errorMsg = e.Message;
                docxml = "error: " + errorMsg;
            }

            if (debug)
                docxml = "error: TESTING ERRORS";

            return docxml;
        }

        public String getActiveDocXml()
        {
            string docxml = "";
            try
            {
                string wpml = Globals.ThisAddIn.Application.ActiveDocument.WordOpenXML;
                docxml = Transform.GetActiveDocumentXml(wpml);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                docxml = "error: " + errorMsg;
            }

            if (debug)
                docxml = "error: TESTING ERRORS";

            return docxml;
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
                stylesxml = "error: "+errorMsg;
            }

            if (debug)
                stylesxml = "error: TESTING ERRORS";
            
            return stylesxml;

        }

        public String getSentenceAtCursor()
        {
            //first get SentenceCount, 
            //next check to see if last thing selected is in a table.
            //if selection exists
            //   for table, return xml for last selected cell, else, return xml for last selected sentence
            //else (no selection)
            //   for table, retun xml for cell cursor is in, else, return xml for sentence cursor is on

            string wpml = "";
            try
            {

               int count = Globals.ThisAddIn.Application.Selection.Range.Sentences.Count;
               Word.Range rng = Globals.ThisAddIn.Application.Selection.Range;
               int stTst = rng.Start;
               int edTst = rng.End;
               string xmlizable = "";

               Word.Table testTbl = null;
               Word.Cell cell = null;
               bool tblExists = false;

               //check for existance of table at cursor
               try
               {
                   testTbl = rng.Tables[1];
                   tblExists = true;
                   int cellCount = Globals.ThisAddIn.Application.Selection.Cells.Count;
                   cell = Globals.ThisAddIn.Application.Selection.Cells[cellCount];
               }
               catch (Exception e)
               {
                   tblExists = false;
               }

                if (stTst < edTst)
                {
                    object startLocation = stTst;
                    object endLocation = edTst;

                    rng = Globals.ThisAddIn.Application.Selection.Range.Sentences[count];
                    rng.Select();

                    if (tblExists)
                    {
                        xmlizable = cell.Range.WordOpenXML;
                        //tables always append empty paragraph; remove and return table only
                        wpml = Transform.ConvertToWPMLFromTextIdx(xmlizable,0);
                    }
                    else
                    {
                        xmlizable = Globals.ThisAddIn.Application.Selection.WordOpenXML;
                        wpml = Transform.ConvertToWPMLFromText(xmlizable);
                    }

                    rng = Globals.ThisAddIn.Application.ActiveDocument.Range(ref startLocation, ref endLocation);
                    rng.Select();
                }
                else
                {
                    int origStart = rng.Start;
                    int origEnd = rng.End;
                    int newStart = origEnd;// -1;
                    int newEnd = origEnd;
                    object startLocation = newStart;
                    object endLocation = newEnd;

                    //need to grab range where cursor is for xml preview
                    rng = Globals.ThisAddIn.Application.Selection.Range.Sentences[1];
                    rng.Select();

                    //check if cursor is in table
                    if (tblExists)
                    {
                        xmlizable = cell.Range.WordOpenXML;
                        //tables always append empty paragraph; remove and return table only
                        wpml = Transform.ConvertToWPMLFromTextIdx(xmlizable, 0);
                    }
                    else
                    {
                        xmlizable = Globals.ThisAddIn.Application.Selection.WordOpenXML;
                        wpml = Transform.ConvertToWPMLFromText(xmlizable);
                    }

                    //return range to original state (selection or cursor position)
                    rng = Globals.ThisAddIn.Application.ActiveDocument.Range(ref startLocation, ref endLocation);
                    rng.Select();

                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                wpml = "error: "+errorMsg;
            }

            if(debugMsg)
               MessageBox.Show("returning wpml: " + wpml);

            if(debug)
               wpml = "error: TESTING ERRORS";

            return wpml;
        }

        //have to remove namespaces to be able to insert; 
        //Office is very particular about where namespaces located in XML.
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

            return tmp;

        }

        public String insertBlockContent(string blockContent, string stylesXml)
        {
            string clean = "";
            clean = removeNamespaces(blockContent);
            string msg = insertBlock(blockContent, stylesXml);

            return msg;

        }

        //insert block level element into document.xml. 
        //If we want to add a style, we have to update styles.xml simultaneously.
        //If we update styles, and there's no reference to the added style in document.xml, Word will
        //consume the xml without error, but not retain the style.  Likewise, we can feed a block
        //level element to Word with a style that doesn't exist in styles.xml.  
        //Word will consume the block without error and assing the block the 
        //default/currently selected style.
        //For future, would like finer grained updates for styles.  Pass block and style defintion, instead of entire styles.xml 
        public String insertBlock(String blockContent, String stylesXml)
        {
            string message = "";
            string wpml = blockContent;
            string newStyle = stylesXml;
            object missing = System.Reflection.Missing.Value;

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
                message = "error: "+errMsg;
            }

            if (debug)
                message = "error: TESTING ERRORS";

            return message;
        }

        public String insertText(string text)
        {
            string message = "";
            try
            {
                Word.Range rng = Globals.ThisAddIn.Application.Selection.Range;
                rng.Text = text;
            }catch(Exception e)
            {
                string errMsg = e.Message;
                message = "error: " + errMsg;
            }

            if (debug)
                message = "error: TESTING ERRORS";

            return message;


        }

        public String getTempPath()
        {
            string tmpPath = "";
            try
            {
                tmpPath = System.IO.Path.GetTempPath();

            }catch(Exception e)
            {
                string errMsg = e.Message;
                tmpPath = "error: " + errMsg;
            }

            return tmpPath;
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

        public String insertImage(string imageuri, string username, string pwd)
        {
            string message="";
            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

                System.Net.WebClient Client = new System.Net.WebClient();
                Client.Credentials = new System.Net.NetworkCredential(username, pwd);
                byte[] bytearray = Client.DownloadData(imageuri);
                Image img = byteArrayToImage(bytearray);

                //backup clipboard
                IDataObject bak = Clipboard.GetDataObject();
                string text = "";
                if (bak.GetDataPresent(DataFormats.Text))
                {
                    text = (String)bak.GetData(DataFormats.Text);
                }
          
                //place on clipboard
                System.Windows.Forms.Clipboard.SetImage(img); 
                Globals.ThisAddIn.Application.Selection.Range.Paste();
                if(!(text.Equals("")))
                   Clipboard.SetText(text);
                


            }catch(Exception e)
            {
                string errMsg = e.Message;
                message = "error: " + errMsg;
            }

            return message;
        }

        public String getSelectionText(int idx, string delimiter)
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
                    wpml = Transform.ExtractTextValuesFromXML(xmlizable, idx, delimiter);
                }
                else
                {
                    wpml = "";
                }
            }


            catch (Exception e)
            {
                string errorMsg = e.Message;
                wpml = "error: " + errorMsg;
            }

            if (debugMsg)
                MessageBox.Show("returning wpml: " + wpml);

            if (debug)
                wpml = "error: Testing errors";

            return wpml;

        }


    }
}
