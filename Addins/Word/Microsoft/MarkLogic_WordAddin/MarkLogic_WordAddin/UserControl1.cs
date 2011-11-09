/*Copyright 2008-2011 Mark Logic Corporation

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
//using DocumentFormat.OpenXml.Packaging;
using System.Web;



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
        private string addinVersion = "2.0-1";  //automate update of this
        HtmlDocument htmlDoc;
        public Word.Document udoc;
      


        private string lastAddedCtrlTitle;
      
        public UserControl1(Word.Document the_doc)
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
                webBrowser1.AllowWebBrowserDrop = false; //false
                webBrowser1.IsWebBrowserContextMenuEnabled = false; //false
                webBrowser1.WebBrowserShortcutsEnabled = true;
                webBrowser1.ObjectForScripting = this;
                webBrowser1.Navigate(webUrl);
                webBrowser1.ScriptErrorsSuppressed = true; 

                this.webBrowser1.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser1_DocumentCompleted);


                udoc = the_doc; // Globals.ThisAddIn.Application.ActiveDocument;
                try
                {
                    udoc.ContentControlOnEnter += new Word.DocumentEvents2_ContentControlOnEnterEventHandler(this.ThisDocument_ContentControlOnEnter);
                    udoc.ContentControlOnExit += new Word.DocumentEvents2_ContentControlOnExitEventHandler(this.ThisDocument_ContentControlOnExit);
                    udoc.ContentControlAfterAdd += new Word.DocumentEvents2_ContentControlAfterAddEventHandler(this.ThisDocument_ContentControlAfterAdd);
                    udoc.ContentControlBeforeDelete += new Word.DocumentEvents2_ContentControlBeforeDeleteEventHandler(this.ThisDocument_ContentControlBeforeDelete);

                    //following only fired when control bound to custom xml part
                    udoc.ContentControlBeforeContentUpdate += new Word.DocumentEvents2_ContentControlBeforeContentUpdateEventHandler(this.ThisDocument_ContentControlBeforeContentUpdate);
                    udoc.ContentControlBeforeStoreUpdate += new Word.DocumentEvents2_ContentControlBeforeStoreUpdateEventHandler(this.ThisDocument_ContentControlBeforeStoreUpdate);

                    //following fires for building block insert, opp for trackin re-use?
                    //different from above, need to think how we'll return
                    //udoc.BuildingBlockInsert += new Word.DocumentEvents2_BuildingBlockInsertEventHandler
                }
                catch (Exception e)
                {
                    MessageBox.Show("error: unable to add Content Control Event handlers to Document. " + e.Message);
                }
                
            }   

        }

        //================= BEGIN CONTENT CONTROL EVENT HANDLERS ====================================
        public void ThisDocument_ContentControlOnEnter(Word.ContentControl contentControl)
        {
            string parentTag = "";
            string parentID = "";

            try
            {
                Word.ContentControl parent = contentControl.ParentContentControl;
                parentTag = parent.Tag;
                parentID = parent.ID;

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }

             contentControlOnEnter(contentControl.ID, contentControl.Tag, contentControl.Title, contentControl.Type.ToString(), contentControl.LockContentControl.ToString(), contentControl.LockContents.ToString(), parentTag, parentID);
        }

        public void ThisDocument_ContentControlOnExit(Word.ContentControl contentControl, ref bool cancel)
        {

            string parentTag = "";
            string parentID = "";

            try
            {
                Word.ContentControl parent = contentControl.ParentContentControl;
                parentTag = parent.Tag;
                parentID = parent.ID;

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }

            contentControlOnExit(contentControl.ID, contentControl.Tag, contentControl.Title, contentControl.Type.ToString(), contentControl.LockContentControl.ToString(), contentControl.LockContents.ToString(), parentTag, parentID);
 
        }

        public void ThisDocument_ContentControlAfterAdd(Word.ContentControl contentControl, bool InUndoRedo)
        {
            string parentTag = "";
            string parentID = "";

            try
            {
                Word.ContentControl parent = contentControl.ParentContentControl;
                parentTag = parent.Tag;
                parentID = parent.ID;

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }
            //MessageBox.Show(contentControl.ID+ contentControl.Tag+ contentControl.Title+ contentControl.Type.ToString()+ contentControl.LockContentControl.ToString()+ contentControl.LockContents.ToString()+ parentTag+ parentID);
            contentControlAfterAdd(contentControl.ID, contentControl.Tag, contentControl.Title, contentControl.Type.ToString(), contentControl.LockContentControl.ToString(), contentControl.LockContents.ToString(), parentTag, parentID);
        }

        public void ThisDocument_ContentControlBeforeDelete(Word.ContentControl contentControl, bool InUndoRedo)
        {
            string parentTag = "";
            string parentID = "";

            try
            {
                Word.ContentControl parent = contentControl.ParentContentControl;
                parentTag = parent.Tag;
                parentID = parent.ID;

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }

            contentControlBeforeDelete(contentControl.ID, contentControl.Tag, contentControl.Title, contentControl.Type.ToString(), contentControl.LockContentControl.ToString(), contentControl.LockContents.ToString(), parentTag, parentID);
        }

        public void ThisDocument_ContentControlBeforeContentUpdate(Word.ContentControl contentControl, ref string content)
        {
            string parentTag = "";
            string parentID = "";

            try
            {
                Word.ContentControl parent = contentControl.ParentContentControl;
                parentTag = parent.Tag;
                parentID = parent.ID;

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }

            contentControlBeforeContentUpdate(contentControl.ID, contentControl.Tag, contentControl.Title, contentControl.Type.ToString(),contentControl.LockContentControl.ToString(), contentControl.LockContents.ToString(), parentTag, parentID);
        }

        public void ThisDocument_ContentControlBeforeStoreUpdate(Word.ContentControl contentControl, ref string content)
        {
            string parentTag = "";
            string parentID = "";

            try
            {
                Word.ContentControl parent = contentControl.ParentContentControl;
                parentTag = parent.Tag;
                parentID = parent.ID;

            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
            }

            contentControlBeforeStoreUpdate(contentControl.ID, contentControl.Tag, contentControl.Title, contentControl.Type.ToString(), contentControl.LockContentControl.ToString(), contentControl.LockContents.ToString(), parentTag, parentID);
        }

        //ERROR HANDLING
        //catch errors where you can recover
        //otherwise always throw
        //(more relevant to JS, keep in mind wrt UserControl)

        public void contentControlOnEnter(string ccID, string ccTag, string ccTitle, string ccType,string ccLockCtrl, string ccLockContents, string ccParentTag, string ccParentID)
        {
            try
            { 
               object result = this.webBrowser1.Document.InvokeScript("contentControlOnEnter", new String[] { ccID, ccTag, ccTitle, ccType, ccLockCtrl, ccLockContents, ccParentTag, ccParentID });
               string res = result.ToString();

               if (res.StartsWith("error"))
               {
                   MessageBox.Show(res);
               }
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show("Error: " + e.Message);
            }
        }

        public void contentControlOnExit(string ccID, string ccTag, string ccTitle, string ccType, string ccLockCtrl, string ccLockContents, string ccParentTag, string ccParentID)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("contentControlOnExit", new String[] { ccID, ccTag, ccTitle, ccType,ccLockCtrl,ccLockContents, ccParentTag, ccParentID });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show(res);
                }
            }
            catch (Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show("Error: " + e.Message);
            }
        }

        public void contentControlAfterAdd(string ccID, string ccTag, string ccTitle, string ccType,string ccLockCtrl, string ccLockContents, string ccParentTag, string ccParentID)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("contentControlAfterAdd", new String[] { ccID, ccTag, ccTitle, ccType, ccLockCtrl, ccLockContents, ccParentTag, ccParentID });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show(res);
                }
            }
            catch(Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show("Error: " + e.Message);
            }
        }

        public void contentControlBeforeDelete(string ccID, string ccTag, string ccTitle, string ccType, string ccLockCtrl, string ccLockContents, string ccParentTag, string ccParentID)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("contentControlBeforeDelete", new String[] { ccID, ccTag, ccTitle, ccType, ccLockCtrl, ccLockContents, ccParentTag, ccParentID });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show(res);
                }
            }
            catch(Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show("Error: " + e.Message);
            }
        }

        public void contentControlBeforeContentUpdate(string ccID, string ccTag, string ccTitle, string ccType, string ccLockCtrl, string ccLockContents, string ccParentTag, string ccParentID)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("contentControlBeforeContentUpdate", new String[] { ccID, ccTag, ccTitle, ccType, ccLockCtrl, ccLockContents, ccParentTag, ccParentID });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show(res);
                }
            }
            catch(Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show("Error: " + e.Message);
            }
        }

        public void contentControlBeforeStoreUpdate(string ccID, string ccTag, string ccTitle, string ccType, string ccLockCtrl, string ccLockContents, string ccParentTag, string ccParentID)
        {
            try
            {
                object result = webBrowser1.Document.InvokeScript("contentControlBeforeStoreUpdate", new String[] { ccID, ccTag, ccTitle, ccType, ccLockCtrl, ccLockContents, ccParentTag, ccParentID });
                string res = result.ToString();

                if (res.StartsWith("error"))
                {
                    MessageBox.Show(res);
                }
            }
            catch(Exception e)
            {
                string donothing_removewarning = e.Message;
                //MessageBox.Show("Error: " + e.Message);
            }
        }

        //=================== END CONTENT CONTROL EVENT HANDLERS ====================================

  /*    trying different things for drag/drop
   *    private void ButtonDown(object sender, HtmlElementEventArgs mea)  //MouseEventArgs mea)
        {
            this.webBrowser1.DoDragDrop("test", DragDropEffects.Copy);
           // MessageBox.Show("TEST");
        }

        private void MouseTest(object sender, MouseEventArgs mea)
        {
          //  this.webBrowser1.DoDragDrop("test", DragDropEffects.Copy);
          //  MessageBox.Show("HERE");
        }
        */
        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
           
            if (webBrowser1.Document != null)
            {
                htmlDoc = webBrowser1.Document;
                htmlDoc.Click += htmlDoc_Click;
              
                //this.webBrowser1.DoDragDrop  MouseDown += new MouseEventHandler(this.MouseTest);
                //revisit for cust and pastte             
                //this.webBrowser1.Document.MouseDown += new HtmlElementEventHandler(this.ButtonDown);
                //htmlDoc.MouseMove += new MouseEventHandler(Element_MouseMove);
                //htmlDoc.MouseDown  += new MouseButtonEventHandler(Element_MouseLeftButtonDown);
                //htmlDoc.MouseUp += new MouseButtonEventHandler(Element_MouseLeftButtonUp);
                
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

        public String setDocumentWordOpenXML(string wpml)
        {
            string docxml = "";
            object missing = System.Reflection.Missing.Value;

            try
            {
                //consider close/open/insert or just open/insert for adding docs
                //Globals.ThisAddIn.Application.Documents.Add(
                //Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument.Close;
                //could be useful for merges, etc.

                //toreset current doc, unfortunately have to select all at this time
    
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                doc.Application.Selection.InsertXML(Transform.setPackageXML(wpml), ref missing);

                object dir = Word.WdCollapseDirection.wdCollapseEnd;
                Word.Range r = doc.Range(ref missing,ref missing);
                r.Collapse(ref dir);
                r.Select();

            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                docxml = "error: " + errorMsg;
            }

            return docxml;
        }

        public String getDocumentWordOpenXML()
        {
            string wpml = "";
            try
            {
               wpml = Globals.ThisAddIn.Application.ActiveDocument.WordOpenXML;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                wpml = "error: " + errorMsg;
            }

            return wpml;

        }

        public String getSelectionWordOpenXML()
        {
            string wpml = "";
            try
            {
                wpml = Globals.ThisAddIn.Application.Selection.WordOpenXML;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                wpml = "error: " + errorMsg;
            }

            return wpml;
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

        public String insertWordOpenXML(string opc_xml)
        {
            string message = "";
            object missing = Type.Missing;

            try
            {
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

                //save file as test
                /* TextWriter tw = new StreamWriter(@"C:\Users\paven\AppData\Local\Temp\MYOPC.xml");
                   tw.WriteLine(opc_xml);
                   tw.Close();
                **/

                //save file as test

                doc.Application.Selection.InsertXML(opc_xml, ref missing);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                MessageBox.Show("ERROR " + message);
            }

            return message;
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
                   string donothing_removewarning = e.Message;
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

        public String getDocumentName()
        {
            string filename = "";
            try
            {
                filename = Globals.ThisAddIn.Application.ActiveDocument.Name;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                filename = "error: " + errorMsg;
            }

            return filename;
        }

        public String getDocumentPath()
        {
            string path = "";
            try
            {
                path = Globals.ThisAddIn.Application.ActiveDocument.Path;
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

        /*    
        public String openDOCXDirect(string docurl, string username, string pwd)
        {
            MessageBox.Show("trying this drama");
            //string path, string title, string url, string user, string pwd
            string message = "";
            try
            {
                //Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;

                System.Net.WebClient Client = new System.Net.WebClient();
                Client.Credentials = new System.Net.NetworkCredential(username, pwd);
                byte[] byteArray = Client.DownloadData(docurl);

                using (MemoryStream mem = new MemoryStream())
                {

                    mem.Write(byteArray, 0, (int)byteArray.Length);
                    using (WordprocessingDocument wordDoc =
                        WordprocessingDocument.Open(mem, true))
                    {

                       
                    }
                  

                    
                    using (FileStream fileStream = new FileStream(@"C:\Test2.docx",System.IO.FileMode.CreateNew))
                    {
                        mem.WriteTo(fileStream);
                    }

                }
            }
            catch (Exception e)
            {
                string errMsg = e.Message;
                message = "error: " + errMsg;
            }

            return message;
        }
        */

        public String openDOCX(string path, string title, string url, string user, string pwd)
        {
            string message = "";
            object missing = Type.Missing;
            string tmpdoc = "";

            try
            {
                tmpdoc = path + title;
               
                downloadFile(url, tmpdoc, user, pwd);
                object filename = tmpdoc;
                object t = true;
                object f = false;
 
                Word.Document d = Globals.ThisAddIn.Application.Documents.Open(ref filename, ref missing, ref f, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref t, ref missing, ref missing, ref missing, ref missing);
                d.Activate();

            }
            catch (Exception e)
            {
                //not always true, need to improve error handling or message or both
                string origmsg = "A document with the name '" + title + "' is already open. You cannot open two documents with the same name, even if the documents are in different \nfolders. To open the second document, either close the document that's currently open, or rename one of the documents.";
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }
            return message;
        }

        public string embedOLE(string path, string title, string url, string user, string pwd)
        {
            string message = "";
            string tmpdoc = "";
            object missing = System.Type.Missing;
            bool proceed = false;

            object left = 60;
            object top = 105;
            object width = 500; 
            object height = 150;


            if (title.EndsWith(".pptx") || title.EndsWith(".pptm") ||
                title.EndsWith(".ppsx") || title.EndsWith(".ppsm") ||
                title.EndsWith(".potx") || title.EndsWith(".potm")                                                 )
            {
                width = 300;
                height = 200;
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
                    object filename = tmpdoc;
                    //defaulting args here.  these could be parameters.
                    //you specify classtype or filename, not both
                    Globals.ThisAddIn.Application.ActiveDocument.Shapes.AddOLEObject(ref missing, ref filename, ref missing, ref missing, ref missing, ref missing, ref missing, ref left, ref top, ref width, ref height, ref missing);
                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message + e.StackTrace;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string saveActiveDocument(string filename, string url, string user, string pwd)
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
            object missing = System.Type.Missing;
            object fname = filename;
            object format = Word.WdSaveFormat.wdFormatDocumentDefault;
            object t = true;

            try
            {
                Word.Document d = Globals.ThisAddIn.Application.ActiveDocument;
                d.SaveAs(ref fname,ref format, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref t, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,ref missing);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }
       
        //begin undocumented - do we want to pursue bookmarks?----------------------------------------//
        public string insertBookmarkText(string bookmark, string text)
        {
            string message = "";

            try
            {
                Word.Bookmarks bs = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks;
                object idx = 1;
                Word.Bookmark b = bs.get_Item(ref idx);
                //MessageBox.Show("Bookmark Name:" + b.Name);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }
        //end undocumented ---------------------------------------------------------------------------//


        //getContentControlIdsByTag
        //getContentControlIdsByTitle
        //use arrow keys to navigate (does doc need to be protected?)
        //also, can map content control to custom part using quick parts

        public string getAllContentControlInfo()
        {
            string message = "";
            string[] ids;
            int i = 0;
            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;
                ids = new string[ccs.Count];
                foreach (Word.ContentControl cc in ccs)
                {
                    ids[i] = cc.ID;
                    i++;


                    //MessageBox.Show("Control Tag:" + cc.Tag + " title: " + cc.Title + " id: "+cc.ID);
                   
                }
            }
            catch (Exception e)
            {
                message = "error: " + e.Message;
            }

            return message;
        }

        public string getContentControlIds()
        {   //id always present, system generated; consistent (unlike customxmlparts); check in XML
            string message = "";
           
            string ids = "";

            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;
                //MessageBox.Show("CONTROL COUNT" + ccs.Count);
                
       
                foreach (Word.ContentControl cc in ccs)
                {
                    //MessageBox.Show("Control Tag:" + cc.Tag + " title: " + cc.Title + " id: "+cc.ID);
                    ids = ids + cc.ID + "|";
                }

                ids=ids.Remove(ids.Length-1);
                message = ids;

        /*        Word.StoryRanges ranges = Globals.ThisAddIn.Application.ActiveDocument.StoryRanges;
                foreach (Word.Range range in ranges)
                {
                    try
                    {
                       // Word.ContentControls allControls = FindContentControls(range);
                       // foreach (Word.ContentControl cc in allControls)
                       // {
                       //     MessageBox.Show(cc.Title + "|" + cc.Tag);
                       // }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("ERROR IS HERE" + e.Message);
                    }
                }

                try
                {
                  //  MessageBox.Show("1 INLINESHAPE COUNT" + Globals.ThisAddIn.Application.ActiveDocument.InlineShapes.Count);
                    foreach (Word.Shape shp in Globals.ThisAddIn.Application.ActiveDocument.InlineShapes)
                    {
                    //     MessageBox.Show("1 INLINESHAPE CC COUNTS" + shp.TextFrame.TextRange.ContentControls.Count);
                       //  Word.ContentControls ccc = shp.TextFrame.TextRange.ContentControls;
                       //  foreach (Word.ContentControl cc in ccc)
                       //  {
                      //       MessageBox.Show(cc.Title + "|" + cc.Tag);
                       //  }


                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("1 ERROR" + e.Message);
                }

                try
                {
                      //MessageBox.Show("2 SHAPE COUNT" + Globals.ThisAddIn.Application.ActiveDocument.Shapes.Count);
                      foreach (Word.Shape shp in Globals.ThisAddIn.Application.ActiveDocument.Shapes)
                      {
                         try
                         {

                             //MessageBox.Show("SHAPE TYPE " + shp.Type + "  2 SHAPE CC COUNTS" + shp.TextFrame.TextRange.ContentControls.Count);
                            // Word.ContentControls ccc = shp.TextFrame.TextRange.ContentControls;
                            // foreach (Word.ContentControl cc in ccc)
                            // {
                              //   MessageBox.Show(cc.Title + "|" + cc.Tag);
                           //  }
                         }
                         catch (Exception e)
                         {
                             MessageBox.Show("not able to have text " + e.Message);
                         }


                      }
                }
                catch (Exception e)
                {
                    MessageBox.Show("2 ERROR" + e.Message);
                }

                try
                {
                    //MessageBox.Show("SHAPE COUNT" + Globals.ThisAddIn.Application.ActiveDocument.CommandBars.Count );
                }
                catch (Exception e)
                {
                    MessageBox.Show("3 ERROR" + e.Message);
                }

                try
                {
                    //MessageBox.Show("SHAPE COUNT" + Globals.ThisAddIn.Application.ActiveDocument.Fields.Count );
                }
                catch (Exception e)
                {
                    MessageBox.Show("4 ERROR" + e.Message);
                }

                try
                {
                    //MessageBox.Show("SHAPE COUNT" + Globals.ThisAddIn.Application.ActiveDocument.FormFields.Count );
                }
                catch (Exception e)
                {
                    MessageBox.Show("5 ERROR" + e.Message);
                }
                try
                {
                    //MessageBox.Show("SHAPE COUNT" + Globals.ThisAddIn.Application.ActiveDocument.Tables.Count );
                }
                catch (Exception e)
                {
                    MessageBox.Show("6 ERROR" + e.Message);
                }

                

           
               // foreach (Word.Shape shp in Globals.ThisAddIn.Application.ActiveDocument.InlineShapes)
               // {
               //     MessageBox.Show("SHAPE COUNTS" + shp.TextFrame.TextRange.ContentControls.Count);
               //     Word.ContentControls ccc = shp.TextFrame.TextRange.ContentControls;
               //     foreach (Word.ContentControl cc in ccc)
               //     {
               //         MessageBox.Show(cc.Title + "|" + cc.Tag);
               //     }


                //}
         * */
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                //MessageBox.Show("MESSAGE"+message);
            }

            return message;
        }

        private static Word.ContentControls FindContentControls(Word.Range rangeStory)
        {
            Word.ContentControls contentControls = null;
           // Word.StoryRanges other = range;
           // foreach (Word.Range range in other)
           // {
                if (rangeStory.StoryType != Word.WdStoryType.wdTextFrameStory)
                {
                    //MessageBox.Show("NOT A TEXT FRAME");

                  
                }

                //Word.Range rangeStory = range;
                //do
               // {
                 //   contentControls = null;
                    try
                    {
                        if (/*rangeStory != null && rangeStory.ShapeRange != null &&*/ rangeStory.ContentControls.Count > 0)
                        {
                           contentControls = rangeStory.ContentControls;
                         
                        }
                    }
                    catch (COMException) { }
                    if (contentControls != null)
                    {
                        //MessageBox.Show("COUNT IN FUNC" + contentControls.Count);
                        return contentControls;
                    }

                   
                   // rangeStory = rangeStory.NextStoryRange;

                //}
                //while (rangeStory != null);
           // }

            return contentControls;
        }

        public string getContentControlIdsByTag(string tag)
        {
            string message = "";
            string ids = "";

            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.SelectContentControlsByTag(tag);
                foreach (Word.ContentControl cc in ccs)
                {
                   ids = ids + cc.ID + "|";
                }

                ids = ids.Remove(ids.Length - 1);
                message = ids;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string getContentControlIdsByTitle(string title)
        {
            string message = "";
            string ids = "";

            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.SelectContentControlsByTitle(title);
                foreach (Word.ContentControl cc in ccs)
                {
                     ids = ids + cc.ID + "|";
                }

                ids = ids.Remove(ids.Length - 1);
                message = ids;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }


        public string getContentControlInfo(string ccid)
        {
            string message = "";
            string info = "";
            string parentTag = "";
            string parentID = "";
            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;
                foreach (Word.ContentControl cc in ccs)
                {
                    if (cc.ID.Equals(ccid))
                    {
                        try
                        {
                            Word.ContentControl parent = cc.ParentContentControl;
                            parentTag = parent.Tag;
                            parentID = parent.ID;

                        }
                        catch (Exception e)
                        {
                            //do nothing, not parent
                            string donothing_removewarning = e.Message;
                        }

                        //info = info + cc.Tag + "|" + cc.Title + "|" + cc.Type +"|" + parentTag +"|"+parentID;;
                        info = info + cc.Tag + "|" + cc.Title + "|" + cc.Type + "|" 
                                    + cc.LockContentControl.ToString() + "|" + cc.LockContents.ToString() + "|"
                                    + parentTag + "|" + parentID; 
                    }
                } 

                message = info;
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string insertContentControlImage(string ccid, string imageuri, string username, string pwd)
        {
            string message = "";

            try
            {
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
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;
                foreach (Word.ContentControl cc in ccs)
                {
                    //if (cc.Tag.Equals(tag))
                    if(cc.ID.Equals(ccid))
                    {
                        cc.Range.Paste();
                    }
                }

                if (!(text.Equals("")))
                    Clipboard.SetText(text);
            }         
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;

        }

        delegate string ICT(string ccid, string text);

        public string insertTextForControl(string ccid, string text)
        {
            string message = "";

            Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;
            foreach (Word.ContentControl cc in ccs)
            {
                //if (cc.Tag.Equals(tag))
                if(cc.ID.Equals(ccid))
                {
                    cc.Range.InsertAfter(text);
                }
            }

            return message;
        }

        public  void Done(IAsyncResult result)
        {
          ICT ict = (ICT)result.AsyncState;
          string isSuccessful = ict.EndInvoke(result);
        }

        public string insertContentControlText(string ccid, string text)
        {
            string message = "";
            try
            {
                ICT ict = insertTextForControl;
                ict.BeginInvoke(ccid, text, Done, ict);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }


    /*    not needed at this time, just lock parent for editing/control removal - similar functionality
     *    -group locks by default, only option to set to not be removed
     *    -can't lock for editing, but in desing mode, this goes away? #fail
     *    -better off using a regular richtext control and locking for now
     * public string groupContentControls(string tag, string controls)
        {
            string message = "";
            string[] tokens = controls.Split('|');
            return message;

        }
     * */
        public string getLastAddedControlTitle()
        {
            return lastAddedCtrlTitle;
        }

        public string addContentControl(string tag, string title, string type, string insertpara, string parent)
        {
            string message = "";
            bool breakflag = false;
            lastAddedCtrlTitle = title;
          
            if (insertpara.ToUpper().Equals("TRUE"))
                  breakflag = true;

            try
            {
                string upperbb = "wdContentControlBuildingBlockGallery".ToUpper();
                string upperpic = "wdContentControlPicture".ToUpper();
                string uppercombo = "wdContentControlComboBox".ToUpper();
                string upperdropdown = "wdContentControlDropdownList".ToUpper();
                string upperdate = "wdContentControlDate".ToUpper();
                string uppergroup = "wdContentControlGroup".ToUpper();
                string uppertext = "wdContentControlText".ToUpper();

                Word.WdContentControlType ccType;

                if (type.ToUpper().Equals(upperpic))
                    ccType = Word.WdContentControlType.wdContentControlPicture;
                else if (type.ToUpper().Equals(uppercombo))
                    ccType = Word.WdContentControlType.wdContentControlComboBox;
                else if (type.ToUpper().Equals(upperdropdown))
                    ccType = Word.WdContentControlType.wdContentControlDropdownList;
                else if (type.ToUpper().Equals(upperdate))
                    ccType = Word.WdContentControlType.wdContentControlDate;
                else if (type.ToUpper().Equals(uppergroup))
                    ccType = Word.WdContentControlType.wdContentControlGroup;
                else if (type.ToUpper().Equals(uppertext))
                    ccType = Word.WdContentControlType.wdContentControlText;
                else if (type.ToUpper().Equals(upperbb))
                    ccType = Word.WdContentControlType.wdContentControlBuildingBlockGallery;
                else
                    ccType = Word.WdContentControlType.wdContentControlRichText;


                object missing = Type.Missing;
                Globals.ThisAddIn.Application.Selection.Range.Select();

                //check based on flag
                if(breakflag)
                  Globals.ThisAddIn.Application.Selection.Range.InsertParagraphAfter();
                else
                  Globals.ThisAddIn.Application.Selection.Range.InsertAfter("  ");

                Globals.ThisAddIn.Application.Selection.Range.Select();

                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                if (!(parent.Equals("")))
                {

                    foreach (Word.ContentControl cc in ccs)
                    {
                        if(cc.ID.Equals(parent))
                        {
                            try
                            {
                                cc.Range.Select();

                                if (breakflag)
                                {
                                    cc.Range.InsertParagraphAfter();
                                }
                                else
                                {
                                    cc.Range.Application.Selection.InsertAfter("  "); 
                                }

                                cc.Range.Select();

                                object startRange = cc.Range.Application.Selection.Range.Characters.Last; // r selection.Range.Characters.First;
                                Word.ContentControl newCC = ccs.Add(ccType, ref startRange);
                                newCC.Range.Text = "";
                                newCC.Tag = tag;
                                newCC.Title = title;

                                message = newCC.ID;
                                break;


                            }
                            catch (Exception e)
                            {
                                string errorMsg = e.Message;
                                message = "error: " + errorMsg;
                            }
                        }
                    }
                }
                else
                {
                    Word.ContentControl myControl = ccs.Add(ccType/*Word.WdContentControlType.wdContentControlRichText*/, ref missing);
                    myControl.Tag = tag;
                    myControl.Title = title;
                    message = myControl.ID;
                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string removeContentControl(string ccid, string deletecontents)
        {
            string message = "";
            object missing = Type.Missing;

            bool delete = false;
            try
            {
                if (deletecontents.ToUpper().Equals("TRUE"))
                    delete = true;

                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    if (cc.ID.Equals(ccid))
                    {
                       

                        //foreach (Word.Table t in cc.Range.Tables)
                          //  t.Delete();

                        cc.Delete(delete);
                        
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

        //could add other functions to handle ContentControlDate* options , ComboBox, etc.

        //could return bool ("true" or "false"), if we passed id instead of tag
        //or do we want a property on SimpleContentControl, or a function to check lockStatus?
        //sets whether or not control can be deleted
        public string lockContentControl(string ccid)
        {
            string message = "";
            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    if (cc.ID.Equals(ccid))
                    {
                        cc.LockContentControl = true;
                    
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

        //could return bool ("true" or "false"), if we passed id instead of tag
        //or do we want a property on SimpleContentControl, or a function to check lockStatus?
        //sets whether or not control can be deleted
        public string unlockContentControl(string ccid)
        {
            string message = "";

            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    if (cc.ID.Equals(ccid))
                        cc.LockContentControl = false;

                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        //could potentially return a bool if we passed id instead of tag
        public string lockContentControlContents(string ccid)
        {
            string message = "";
            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    if (cc.ID.Equals(ccid))
                        cc.LockContents = true;

                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        //could potentially return a bool if we passed id instead of tag
        public string unlockContentControlContents(string ccid)
        {
            string message = "";
            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    if (cc.ID.Equals(ccid))
                        cc.LockContents = false;

                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        //void, could return bool, change to id(vs. tag)?
        public string mapContentControl(string ccid, string xpath, string prefix, string cid)
        {
            string message = "";
            bool mapped = false;

            try
            {
                Office.CustomXMLPart mypart = Globals.ThisAddIn.Application.ActiveDocument.CustomXMLParts.SelectByID(cid);
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    //if (cc.Tag.Equals(tag/*"pttesttag"*/))
                    if(cc.ID.Equals(ccid))
                    { 
                        mapped = cc.XMLMapping.SetMapping(xpath, prefix, mypart);
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

        public string getContentControlDropDownListEntrySelectedText(string ccid)
        {
            string message = "";

            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    if (cc.ID.Equals(ccid))
                    {
                        message = cc.Range.Text;
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

        public string getContentControlDropDownListEntrySelectedValue(string ccid)
        {
            string message = "";

            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    if (cc.ID.Equals(ccid))
                    {
                        string displaytext =  cc.Range.Text;
                        
                        Word.ContentControlListEntries ccles = cc.DropdownListEntries;
                        foreach (Word.ContentControlListEntry ccle in ccles)
                        {
                            if (ccle.Text.Equals(displaytext))
                                message = ccle.Value;

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

        public string addContentControlDropDownListEntries(string ccid, string text, string value, string index)
        {
            string message = "";
            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    //if (cc.Tag.Equals(tag))
                    if(cc.ID.Equals(ccid))
                    {
                        cc.DropdownListEntries.Add(text, value, Int32.Parse(index));

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

        public string getParentContentControlInfo()
        {
            string message = "";
            string info = "";
            string parentTag = "";
            string parentID = "";

            Word.Range current = Globals.ThisAddIn.Application.Selection.Range;
            try
            {
                Word.ContentControl cc = current.ParentContentControl;
                info = info +cc.ID+"|"+ cc.Tag + "|" + cc.Title + "|" + cc.Type + "|" 
                            + cc.LockContentControl.ToString() + "|" + cc.LockContents.ToString();

                try
                {
                    Word.ContentControl parent = cc.ParentContentControl;
                    parentTag = parent.Tag;
                    parentID = parent.ID;
                    info = info + "|" + parentTag + "|" + parentID;
                }catch(Exception e)
                {
                    string donothing_removewarning = e.Message;
                    info = info + "|" + parentTag + "|" + parentID;;

                }

            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                info = "error: " + errorMsg;
            }

            message = info;
            return message;
        }


        public string setCursorAfterContentControl(string ccid)
        {          
            string message = "";
            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    //if (cc.Tag.Equals(tag))
                    if(cc.ID.Equals(ccid))
                    {
                        //object dir = Word.WdCollapseDirection.wdCollapseEnd;
                        
                        Word.Range r = cc.Range;
                        //r.Collapse(ref dir);
                        int idx = r.End;
                        int charstartrange = idx +1;
                        object unit = Word.WdUnits.wdCharacter;
                        object count = charstartrange;
                        r.MoveStart(ref unit, ref count);
                        r.End = charstartrange;
                        r.Select();

                     
                    }
                }
            }catch(Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                //MessageBox.Show(message);
            }

            return message;
                    
        }

        public string deleteSelection()
        {
            string message = "";

            try
            {
                object missing = Type.Missing;
                Globals.ThisAddIn.Application.Selection.Range.Delete(ref missing, ref missing);
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                //MessageBox.Show(message);
            }

            return message;
        }

        public string setCursorBeforeContentControl(string ccid)
        {
            string message = "";
            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    //if (cc.Tag.Equals(tag))
                    if (cc.ID.Equals(ccid))
                    {
                        //object dir = Word.WdCollapseDirection.wdCollapseEnd;

                        Word.Range r = cc.Range;
                        //r.Collapse(ref dir);
                        int idx = r.Start;
                        int charstartrange = idx - 1;
                        object unit = Word.WdUnits.wdCharacter;
                        object count = charstartrange;
                        r.MoveStart(ref unit, ref count);
                        r.End = charstartrange;
                        r.Select();


                    }
                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                //MessageBox.Show(message);
            }

            return message;

        }

        public string setContentControlFocus(string ccid)
        {
            string message = "";
            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    //if (cc.Tag.Equals(tag))
                    if (cc.ID.Equals(ccid))
                    {
                        object dir = Word.WdCollapseDirection.wdCollapseStart;

                        Word.Range r = cc.Range;
                        r.Collapse(ref dir);

                        r.Select();
                    }
                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
                //MessageBox.Show(message);
            }

            return message;

        }
    
        public string setContentControlPlaceholderText(string ccid, string pltext,string cleartext)
        {
             string message = "";
             bool clear = false;

             if (cleartext.ToUpper().Equals("TRUE"))
             {
                  clear = true;
             }

             try
             {
                 Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;
                 
                 foreach (Word.ContentControl cc in ccs)
                 {
                     if (cc.ID.Equals(ccid))
                     {
                         object missing = Type.Missing;

                         if (clear)
                         {
                             //cc.Range.Delete(ref missing, ref missing);
                             cc.Range.Text = "";
                         }

                         cc.SetPlaceholderText(null, null, pltext);
                         
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

        public string getContentControlText(string ccid)
        {         
             string message = "";
             try
             {
                 Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                 foreach (Word.ContentControl cc in ccs)
                 {
                     if (cc.ID.Equals(ccid))
                     {
                         message = cc.Range.Text;
                        
                     }
                 }

             }catch (Exception e)
             {
                 string errorMsg = e.Message;
                 message = "error: " + errorMsg;
             }

             return message;
        }

        public string getContentControlWordOpenXML(string ccid)
        {
            string message = "";

            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    if (cc.ID.Equals(ccid))
                    {
                        message = cc.Range.WordOpenXML;

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

        public string setContentControlTag(string ccid, string tag)
        {
            string message = "";

            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    if (cc.ID.Equals(ccid))
                    {
                        cc.Tag = tag;
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

        public string setContentControlTitle(string ccid, string title)
        {
            string message = "";

            try
            { 
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    if (cc.ID.Equals(ccid))
                    {
                        cc.Title = title;
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

        public string setContentControlStyle(string ccid, string style)
        {
            string message = "";
            object styleHeading2 = style;
          //  object styleHeading3 = "Heading 3";


            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    if(cc.ID.Equals(ccid))
                       cc.Range.set_Style(ref styleHeading2);
 
                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string hideContentControlRange(string ccid)
        {
            string message = "";

            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    if (cc.ID.Equals(ccid))
                        cc.Range.Font.Hidden = 1;

                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        public string displayContentControlRange(string ccid)
        {
            string message = "";

            try
            {
                Word.ContentControls ccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;

                foreach (Word.ContentControl cc in ccs)
                {
                    if (cc.ID.Equals(ccid))
                        cc.Range.Font.Hidden = 0;

                }
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
            }

            return message;
        }

        /**
         * Starts compare with the current document and the document retrieved from
         * the specified uri using the specified credentials. The file is assumed to 
         * be .xml format.
         */
        //maybe pass filename
        public string mergeWithActiveDocument(string opc_xml)
        {
            string message = "";
           // object oname = "TEMP";
            object f = false;
            object t = true;
            object missing = Type.Missing;
            //object format = Word.WdSaveFormat.wdFormatDocumentDefault;
            //object fname = getTempPath() + "TEMP.docx";
            Microsoft.Office.Interop.Word.Document doc1 = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Interop.Word.Document doc = null;

            try
            {
                
                doc = Globals.ThisAddIn.Application.Documents.Add(ref missing, ref missing, ref missing, ref f);
                doc.Content.Select();
                doc.Application.Selection.InsertXML(opc_xml, ref missing);

            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message = "error: " + errorMsg;
               // MessageBox.Show("error inserting XML: " + e.Message);
            }

            try
            {
                Word.WdCompareDestination dest = Word.WdCompareDestination.wdCompareDestinationNew;// wdCompareDestinationNew; // .wdCompareDestinationNew;   //.wdCompareDestinationNew;
                Word.WdGranularity level = Word.WdGranularity.wdGranularityWordLevel;
                //Word.Document doc1 = Globals.ThisAddIn.Application.ActiveDocument;
                Globals.ThisAddIn.Application.MergeDocuments(doc1, doc, dest,
                                                             level, true, true, true, true, true,
                                                             true,
                                                             true, true, true, true, "", "",
                                                             Word.WdMergeFormatFrom.wdMergeFormatFromOriginal);
                
                //Globals.ThisAddIn.Application.CompareDocuments(doc1, doc, dest, level, true, true, true, true, true, true, true, true, true, true, "", true);


                Globals.ThisAddIn.MERGEFLAG = false;  
                doc.Close(ref f, ref missing, ref missing);

                Globals.Ribbons.Ribbon1.viewTaskPaneButton.Checked = false;
                Globals.ThisAddIn.RemoveAllTaskPanes();
               
            }
            catch (Exception e)
            {
                string errorMsg = e.Message;
                message += "error: " + errorMsg;
                //MessageBox.Show("error in MERGE: " + e.Message);
            }

            return message;
        }

        private void webBrowser1_DocumentCompleted_1(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        /* Review, added these 3 for Jay.
         1)	MLA.setCursorBeforeContentControl(controlID); 
         2)	MLA.setCursorAfterContentControl(controlID);
         3)	MLA.deleteSelection();
        */

       /*
       //a bizzare experiment that actually works, still serializes as 2007 xml on save
       public string insert2003XML()
       {
           string message = "";

           string xml = 
             "<w:document xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml'>" +  // xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>" +
               "<w:body>" +
                 "<w:p>" +
                   "<w:pPr>" +
                   "<w:rPr>" +
                     "<w:u w:val='single'/>" +
                   "</w:rPr>" +
                   "</w:pPr>" + 
                   "<w:r>" +
                     "<w:rPr>" +
                          "<w:u w:val='single'/>" +
                     "</w:rPr>" +
                     "<w:t>TEST UNDERLINE.</w:t>" +
                   "</w:r>" +
                 "</w:p>"+
               "</w:body>" +
             "</w:document>" ;

           try
           {
               object missing = Type.Missing;
               Word.Range r = Globals.ThisAddIn.Application.Selection.Range;
               r.InsertXML(xml, ref missing);
           }
           catch (Exception e)
           {
               MessageBox.Show("ERROR: " + e.Message);
           }
           return message;
       }

       //this again - another experiment
       public string testDragDrop()
       {
           string message = "";
           //this.MouseDown += new MouseEventHandler(this.webBrowser1);

           IDataObject bak = Clipboard.GetDataObject();
           string text = "";
           if (bak.GetDataPresent(DataFormats.Text))
           {
               text = (String)bak.GetData(DataFormats.Text);
           }
           object fo = "t";
           //DataObject dragData = new DataObject(typeof(string), PetesMethodOfDoom(mqr));
           //DataObject dragData = new DataObject(typeof(string), PetesMethodOfDoom(mqr));

           //DragDrop.DoDragDrop(TestBox, dragData, DragDropEffects.All);
           
           //System.Windows.DragDrop.DoDragDrop(fo, "test", System.Windows.DragDropEffects.Copy);
           //webBrowser1.DoDragDrop("test", DragDropEffects.All);
           return message;
       }
        
       */

    }
}
