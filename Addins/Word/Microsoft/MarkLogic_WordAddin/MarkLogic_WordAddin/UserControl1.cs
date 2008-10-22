/*Copyright 2002-2008 Mark Logic Corporation.  All Rights Reserved*/
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


namespace MarkLogic_WordAddin
{   
    [ComVisible(true)]
  //  [ClassInterfaceAttribute(ClassInterfaceType.AutoDispatch)]
  //  [DockingAttribute(DockingBehavior.AutoDock)]
  //  [PermissionSetAttribute(SecurityAction.InheritanceDemand, Name = "FullTrust")]
  //  [PermissionSetAttribute(SecurityAction.LinkDemand, Name = "FullTrust")]

    public partial class UserControl1 : UserControl
    {
        private AddinConfiguration ac = AddinConfiguration.GetInstance();
        private string webUrl = "";
        private bool debug = false;
        private bool debugMsg = false;
        private string color = "";
        private string addinVersion = "@MAJOR_VERSION.@MINOR_VERSION@PATCH_VERSION";


        public UserControl1()
        {
            InitializeComponent();
           // AddinConfiguration ac = AddinConfiguration.GetInstance();
            //bool regEntryExists = checkUrlInRegistry();
            webUrl = ac.getWebURL();
            //MessageBox.Show(webUrl);

            if (webUrl.Equals(""))
            {
                //MessageBox.Show("Unable to find configuration info. Please insure OfficeProperties.txt exists in your system temp directory.  If problems persist, please contact your system administrator.");
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

            }


        }
/*
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
        */
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
                custompiecexml = "error: "+errorMsg;
            }

            if (debug)
                custompiecexml = "error: TESTING ERRORS";

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
                newid = "error: "+errorMsg;
            }
            if (debug)
                newid = "error: testing Errors";

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
                message = "error: "+errorMsg;
            }

            if (debug)
                message = "error: Testing errors.";

            return message;
             
        }

        //currently no way to replace without delete,add, get new id
/*
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
                wpml = "error: "+errorMsg;
            }
            
            if(debugMsg)
               MessageBox.Show("returning wpml: " + wpml);

            if (debug)
                wpml = "error: Testing errors";

            return wpml;

        }
*/
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

        public String replaceActiveDocumentXml(string wpml)
        {
            string docxml = "";
            object missing = System.Reflection.Missing.Value;
            //MessageBox.Show(wpml);

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

        //returns the style for the current block
        public String getSentenceAtCursor()
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
                wpml = "error: "+errorMsg;
            }

            if(debugMsg)
               MessageBox.Show("returning wpml: " + wpml);

            if(debug)
               wpml = "error: TESTING ERRORS";

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
                message = "error: "+errMsg;
            }

            if (debug)
                message = "error: TESTING ERRORS";

            return message;
        }


        //FOLLOWING ARE UNDOCUMENTED AND NOT OFFICIALLY SANCTIONED METHODS
        //THESE MAY BE REMOVED, OR CHANGE
        //====================================================================
        public String getRangeForSelection()
        {
            string message="";
            try
            {

                Word.Range rng = Globals.ThisAddIn.Application.Selection.Range;
                int stTst = rng.Start;
                int edTst = rng.End;

                if (stTst < edTst)
                    message = stTst + ":" + edTst;
            }
            catch (Exception e)
            {
                string errMsg = e.Message; 
                message = "error: " + errMsg;
            }

            if (debug)
                message = "error: TESTING ERRORS";

            return message;
        }

        public String getRangesForTerm(string searchTerm)
        {
            string message = "";
            object searchText = searchTerm;
            object missing = System.Reflection.Missing.Value;
            object start = 0;

            Word.Range rng = Globals.ThisAddIn.Application.ActiveDocument.Content;
            try
            {
                rng.Find.ClearFormatting();
                rng.Find.Forward = true;
                rng.Find.MatchWholeWord = true;   //parameterize (as well as case?sounds like?)   
                rng.Find.Text = searchTerm;

                rng.Find.Execute(
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);

                while (rng.Find.Found)
                {

                    MessageBox.Show("RANGE" + rng.Start + " " + rng.End);
                    message = message + rng.Start + ":" + rng.End + " ";

                        rng.Find.Execute(
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing);
                }
                message=message.Trim();
            }
            catch (Exception e)
            {
                MessageBox.Show("IN ERROR");
                string errMsg = e.Message;
                message = "error: " + errMsg;
            }

            if (debug)
                message = "error: TESTING ERRORS";

            //MessageBox.Show("RETURNING: " + message);
            return message;
        }


        public String addContentControlToRange(string ranges, string title, string tag, bool lockStatus)
        {
            string message = "";
            object missing = System.Reflection.Missing.Value;
            ranges = ranges.Trim();
            char[] delimiter = { ' ' };

            if (!(ranges.Equals("") || ranges == null))
            {
                string[] tmp1 = ranges.Split(delimiter);
                if (tmp1.Length > 0)
                {
                    char[] delimiter2 = { ':' };
                    int count = 0;
                    try
                    {
                        foreach (string x in tmp1)
                        {
                            string[] tmp2 = x.Split(delimiter2);
                            int start = System.Convert.ToInt32(tmp2[0], 10);
                            int end = System.Convert.ToInt32(tmp2[1], 10);
                            object s = start + count;
                            object e = end + count;

                            //count++; //range shifts as each comment added (i know, wtf?)
                            Word.Range rng = Globals.ThisAddIn.Application.ActiveDocument.Range(ref s, ref e);

                            //can't add comment to plain text control, or to control that's locked
                            //how bout adding other control???
                            //which leads to question - unlock, add comment, relock???
                            bool locked = false;
                            bool plainText = false;
                            Word.Range rngSection = rng.Sections[1].Range;
                            Word.ContentControls cs = rngSection.ContentControls;

                            foreach (Word.ContentControl cc in cs)
                            {
                                // MessageBox.Show("TYPE IS" + cc.Type);
                                if (rng.InRange(cc.Range) && locked == false && plainText == false)
                                {
                                    if (cc.LockContents == true)
                                        locked = true;

                                    if (cc.Type.Equals(Microsoft.Office.Interop.Word.WdContentControlType.wdContentControlText))
                                        plainText = true;
                                }
                            }
                            //added if around comments.add line
                            if (!locked && !plainText)
                            {
                                //rng.Comments.Add(rng, ref cmt);

                                Word.ContentControl cControl = rng.ContentControls.Add(Microsoft.Office.Interop.Word.WdContentControlType.wdContentControlRichText, ref missing);
                                cControl.Title = title;
                                cControl.Tag = tag;
                                string t = rng.Text;
                                cControl.Range.Text = t;
                                cControl.LockContents = lockStatus;
                                cControl.LockContents = lockStatus;

                                int len = t.Length + 1;

                                //Have to shift range manually to get out of new Content Control
                                rng.Start = rng.End + len;
                                //have to add 2 to range for controls, 1 to range for comments
                                //in this if as we only increase if we add a control
                                count++;
                                count++;
                            }
                            // MessageBox.Show("VALUE OF start" + start + "end " + end);
                            // MessageBox.Show("TRUE" + rng.StoryLength + " " + rng.StoryType + "");
                        }
                    }
                    catch (Exception ex)
                    {
                        string errMsg = ex.Message;
                        message = "error: " + errMsg;
                    }


                    if (debug)
                        message = "error: TESTING ERRORS";


                }
            }
            return message;

        }

        //can't add comment to locked content controls - change this? (we can make the change)
        //noting for discussion later
        public String addCommentToRange(string ranges, string comment)
        {
            string message = "";
            ranges = ranges.Trim();
            //MessageBox.Show("IN FUNCTION  ranges:"  +ranges + " comment: "+comment);
            char[] delimiter = { ' ' };
            if (!(ranges.Equals("") || ranges == null))
            {

              string[] tmp1 = ranges.Split(delimiter);
              if (tmp1.Length > 0)
              {
                char[] delimiter2 = { ':' };
                int count = 0;
                try
                {
                    foreach (string x in tmp1)
                    {
                        string[] tmp2 = x.Split(delimiter2);
                        int start = System.Convert.ToInt32(tmp2[0], 10);
                        int end = System.Convert.ToInt32(tmp2[1], 10);
                        object s = start + count;
                        object e = end + count;

                       //count++; //range shifts as each comment added (i know, wtf?)
                        Word.Range rng = Globals.ThisAddIn.Application.ActiveDocument.Range(ref s, ref e);

                        object cmt = comment;

                        //added from here to next comment to check for rang in control with lock
                        //can't add comment to plain text control, or to control that's locked
                        //which leads to question - unlock, add comment, relock???
                        bool locked = false;
                        bool plainText = false;
                        Word.Range rngSection = rng.Sections[1].Range;
                        Word.ContentControls cs = rngSection.ContentControls;
                       
                        foreach (Word.ContentControl cc in cs)
                        {
                            // MessageBox.Show("TYPE IS" + cc.Type);
                            if (rng.InRange(cc.Range) && locked == false && plainText == false)
                            {
                                if (cc.LockContents == true)
                                    locked = true;

                                if (cc.Type.Equals(Microsoft.Office.Interop.Word.WdContentControlType.wdContentControlText))
                                    plainText = true;
                            }
                        }
                        //added if around comments.add line
                        if (!locked && !plainText)
                        {
                            rng.Comments.Add(rng, ref cmt);
                            count++;
                        }
                        // MessageBox.Show("VALUE OF start" + start + "end " + end);
                        // MessageBox.Show("TRUE" + rng.StoryLength + " " + rng.StoryType + "");
                    }
                }
                catch (Exception ex)
                {
                    string errMsg = ex.Message;
                    message = "error: " + errMsg;
                }
             

                    if (debug)
                        message = "error: TESTING ERRORS";

                    
              }
            }
            return message;
        }

        //orig, broke this up into 2 functions, one to get ranges, on too add comment based on ranges
        public String addCommentForText(string searchTerm, string comment)//(string comment, string text)
        {
           
            string message = "";
            object searchText = searchTerm;
            object commentText = comment;
            object missing = System.Reflection.Missing.Value;

            object start = 0;

            Word.Range rng = Globals.ThisAddIn.Application.ActiveDocument.Content;
            try
            {
                rng.Find.ClearFormatting();
                rng.Find.Forward = true;
                rng.Find.MatchWholeWord = true;   //parameterize (as well as case?sounds like?)   
                rng.Find.Text = searchTerm;

                rng.Find.Execute(
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);

                while (rng.Find.Found)
                {
                    bool locked = false;
                    bool plainText = false;
                    Word.Range rngSection = rng.Sections[1].Range;
                    Word.ContentControls cs = rngSection.ContentControls;
                    
                    foreach (Word.ContentControl cc in cs)
                    {
                      // MessageBox.Show("TYPE IS" + cc.Type);
                        if (rng.InRange(cc.Range) && locked == false && plainText ==false)
                        {
                            if (cc.LockContents == true)
                                locked = true;

                            if (cc.Type.Equals(Microsoft.Office.Interop.Word.WdContentControlType.wdContentControlText))
                                plainText = true;
                        }
                    }

                    //MessageBox.Show("LOCKED: " + locked + "  PLAINTEXT: " + plainText);

                    if (!locked && !plainText)
                    {
                        try
                        {
                            rng.Comments.Add(rng, ref commentText);
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show("unable to add comment" + e.Message);
                        }
                    }

                    rng.Find.Execute(
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing);
                }
            }catch(Exception e)
            {
                string errMsg = e.Message;
                message = "error: " + errMsg;
            }

            if (debug)
                message = "error: TESTING ERRORS";
         
            return message;

        }


        //will only add control to text that is not in locked control already
        //we can change this, not even sure how useful this function is though
        //noting here for later
        public String addContentControlForText(string searchTerm, string title, string tag, bool lockstatus)//(string comment, string text)
        {
            string message = "";
            object missing = System.Reflection.Missing.Value;
            object start = 0;
            Word.Range rng = Globals.ThisAddIn.Application.ActiveDocument.Content;
           
            try
            {
                rng.Find.ClearFormatting();
                rng.Find.Forward = true;
                rng.Find.MatchWholeWord = true;   //parameterize (as well as case?sounds like?)   
                rng.Find.Text = searchTerm;

                rng.Find.Execute(
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);

                while (rng.Find.Found)
                {
                    bool identical = false;
                    bool found = false;
                    bool lockedcontent = false;
                    bool plainText = false;

                    string foundtag = "";
                    string foundtitle = "";
                    Word.Range rngSection = rng.Sections[1].Range;
                    Word.ContentControls cs = rngSection.ContentControls;


                   // MessageBox.Show("COUNT" + cs.Count);

                    foreach (Word.ContentControl cc in cs)
                    {

                        if (rng.InRange(cc.Range) && found==false  && identical==false && lockedcontent ==false  && plainText==false)
                        {
                            found = true;
                            foundtag = cc.Tag;
                            foundtitle = cc.Title;
                           // MessageBox.Show("CONTROL TAG: " + cc.Tag + "CONTROL TITLE: " + cc.Title);
                          //  MessageBox.Show("TAG: " + tag + "TITLE: " + title);
                            if (foundtag.Equals(tag) && foundtitle.Equals(title))
                            {
                             //   MessageBox.Show("THEYRE EQUAL");
                                identical = true;
                            }
                            else
                            {
                                //MessageBox.Show("They're not equal");
                                found = false;
                            }

                            if (cc.LockContents.Equals(true))
                            {
                                lockedcontent = true;
                            }

                            if (cc.Type.Equals(Microsoft.Office.Interop.Word.WdContentControlType.wdContentControlRichText))
                                plainText = true;



                        }

                    }
                    
                    //the control doesn't exist and is not identical to the existing control
                    if (!found && !identical  && !lockedcontent  && !plainText)
                    {
                        Word.ContentControl cControl = rng.ContentControls.Add(Microsoft.Office.Interop.Word.WdContentControlType.wdContentControlRichText, ref missing);
                        cControl.Title = title;
                        cControl.Tag = tag;
                        MessageBox.Show("LOCKSTATUS" + lockstatus);
                        string t = rng.Text;
                        cControl.Range.Text = t;
                        cControl.LockContentControl = lockstatus;
                        cControl.LockContents = lockstatus;


                        int len = t.Length + 1;

                        //Have to shift range manually to get out of new Content Control
                        rng.Start = rng.End + len;
                        //MessageBox.Show("Adding controls");
                    }
                    else
                    {
                        //MessageBox.Show("Already in control");
                    }

                    rng.Find.Execute(
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing);
                }
            }
            catch (Exception e)
            {
                string errMsg = e.Message;
                message = "error: " + errMsg;
                MessageBox.Show(errMsg);
            }

            if (debug)
                message = "error: TESTING ERRORS";

            return message;

        }

        //have to think about this one, if you delete, you remove the content
        //but what if this control is child of another
        //want to remove control, but not content if parent is content control
        //need to update this to remove control, leaving parents
        //noting for discussion later

        public String deleteContentControl()
        {
            MessageBox.Show("IN FUNCTION");
            string message = "";
            string tag = "5TAGFORCONTROL";
            Word.ContentControls wccs = Globals.ThisAddIn.Application.ActiveDocument.ContentControls;
            Word.Document wd = Globals.ThisAddIn.Application.ActiveDocument;
          

            foreach(Word.ContentControl cc in wccs)
            {
                MessageBox.Show("TAG IS: " + cc.Tag);
                if (cc.Tag.Equals(tag))
                {
                    cc.LockContentControl = false;
                    cc.LockContents = false;
                    cc.Delete(true);
                 
                    MessageBox.Show("REMOVING CONTROL");

                }
            }

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

        public String getSelectionText()
        {
            string message = "";
            object missing = System.Reflection.Missing.Value;
            Word.Selection selection = Globals.ThisAddIn.Application.Selection;//((Word.Window)control.Context).Application.Selection;
            int selectionLength = selection.Range.End - selection.Range.Start;

            if (selectionLength > 0)
            {

                message = selection.Text;
            }

            return message;

        }

        public String insertTextInControl(string text, string tag, bool lockStatus)
        {
            string message = "";
            object missing = System.Reflection.Missing.Value;

            try
            {
                Word.Range rng = Globals.ThisAddIn.Application.Selection.Range;
                //rng.Text = text;
                Word.ContentControl cControl = rng.ContentControls.Add(Microsoft.Office.Interop.Word.WdContentControlType.wdContentControlRichText, ref missing);
                cControl.Title = tag;
                cControl.Tag = tag;
                cControl.Range.Text = text;
                cControl.LockContentControl = lockStatus;
                cControl.LockContents = lockStatus;

            }
            catch (Exception e)
            {
                string errMsg = e.Message;
                message = "error: " + errMsg;
            }

            if (debug)
                message = "error: TESTING ERRORS";

            return message;

        }


        //issue?  seem to lose comments if you add control around commented text
        public String addContentControlToSelection(string tagName, bool lockStatus)
        {
            string message = "";
            object missing = System.Reflection.Missing.Value;
            Word.Selection selection = Globals.ThisAddIn.Application.Selection;//((Word.Window)control.Context).Application.Selection;
            int selectionLength = selection.Range.End - selection.Range.Start;

            if (selectionLength > 0)
            {

                if (CreateContentControlInSelection(
                    selection,
                    tagName,
                    tagName,
                    lockStatus) != null)
                    //DateTime.Now.ToString()) != null)
                {
                    object oUnit = Word.WdUnits.wdCharacter;
                    object oCount = selectionLength + 1;

                    //Gets out of the content control
                   
                    selection.MoveRight(ref oUnit, ref oCount, ref missing);
                }
            }
            

            return message;
        }

        //currently limited to 1 paragraph
        //need to examine plain text vs. rich text controls
        private Word.ContentControl CreateContentControlInSelection(
            Word.Selection selection, string controlName, string controlTag, bool lockStatus)
        {
            Word.ContentControl contentControl = null;

           
            if (selection.Paragraphs.Count > 1)
            {
                MessageBox.Show(
                    "Content control cannot be inserted around multiple paragraphs.",
                    "Microsoft Office Word", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                object oRng = (object)selection.Range;
                contentControl = selection.Range.ContentControls.Add(
                    Word.WdContentControlType.wdContentControlRichText, ref oRng);
                if (contentControl != null)
                {
                    contentControl.Range.Text = selection.Text.Replace("\r", "");
                    contentControl.Tag = controlTag;
                    contentControl.Title = controlName;
                    contentControl.LockContents = lockStatus;
                    contentControl.LockContents = lockStatus;
                }
            }

            return contentControl;
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
 */

    }
}
