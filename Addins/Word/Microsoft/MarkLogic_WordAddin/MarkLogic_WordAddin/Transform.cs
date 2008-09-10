/*Copyright 2002-2008 Mark Logic Corporation.  All Rights Reserved*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace MarkLogic_WordAddin
{

    class Transform
    {
        static XmlNamespaceManager NamespaceManager = null;
        const string bodyXPath = "//pkg:part[@pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml']/pkg:xmlData/w:document/w:body";
        const string paraXPath = "//pkg:part[@pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml']/pkg:xmlData/w:document/w:body/w:p";
        const string documentXPath = "//pkg:part[@pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml']/pkg:xmlData";
        const string stylesXPath = "//pkg:part[@pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml']/pkg:xmlData";
        const string numbersXPath = "//pkg:part[@pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml']/pkg:xmlData";
        const string themesXPath = "//pkg:part[@pkg:contentType='application/vnd.openxmlformats-officedocument.theme+xml']/pkg:xmlData";
        const string paraPropsXPath = "/w:p/w:pPr";
        const string runPropsXPath = "/w:p/w:r/w:rPr";


        static Transform()
        {
            NamespaceManager = new XmlNamespaceManager(
                new NameTable());
            NamespaceManager.AddNamespace(
                "pkg", "http://schemas.microsoft.com/office/2006/xmlPackage");
            NamespaceManager.AddNamespace(
                "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        }

        public static string GetActiveDocXML()
        {
            //Get current Selected Range
            Word.Range tmprng = Globals.ThisAddIn.Application.Selection.Range;
            int tmpRangeStart = tmprng.Start;
            int tmpRangeEnd = tmprng.End;


            //Get docxml for current doc, which will shift range to entire doc
            Globals.ThisAddIn.Application.ActiveDocument.Content.Select();
            string docxml = Globals.ThisAddIn.Application.Selection.WordOpenXML;

            //reset range
            object tmpstart = tmpRangeStart;
            object tmpend = tmpRangeEnd;
            Word.Range rng = Globals.ThisAddIn.Application.ActiveDocument.Range(ref tmpstart, ref tmpend);
            rng.Select();

            return docxml;
        }

        public static string GetStylesXmlFromCurrentDoc(string wordprocessingML)
        {
            StringBuilder response = new StringBuilder();
            try
            {
                XmlDocument document = new XmlDocument();
                document.LoadXml(wordprocessingML);
                XmlNode content = document.SelectSingleNode(stylesXPath, NamespaceManager);
               
                using (XmlTextWriter writer = new XmlTextWriter(
                    new StringWriter(response)))
                {
                    writer.Formatting = Formatting.Indented;
                    content.WriteContentTo(writer);
                }
            }
            catch (Exception e)
            {
                response.Append("Error " + e.Message);
            }
            return response.ToString();
        }

        public static string ConvertToWPMLFromText(string wordprocessingML)
        {
            StringBuilder response = new StringBuilder();
            try
            {
                XmlDocument document = new XmlDocument();
                document.LoadXml(wordprocessingML);
                XmlNode content = document.SelectSingleNode(bodyXPath, NamespaceManager);
                if (content.ChildNodes.Count > 0)
                {
                    content.RemoveChild(content.LastChild);
                }
                using (XmlTextWriter writer = new XmlTextWriter(
                    new StringWriter(response)))
                {
                    writer.Formatting = Formatting.Indented;
                    content.WriteContentTo(writer);
                }
            }
            catch (Exception e)
            {
                response.Append("Error " + e.Message);
            }
            return response.ToString();
        }
       
        public static string ConvertToWPMLFromTextFinalNode(string wordprocessingML)
        {
            StringBuilder response = new StringBuilder();
            try
            {
                XmlDocument document = new XmlDocument();
                document.LoadXml(wordprocessingML);
                XmlNode content = document.SelectSingleNode(bodyXPath, NamespaceManager);
                if (content.ChildNodes.Count > 0)
                {
                    content.RemoveChild(content.LastChild);
                }

                int length = 0;
                int counter = 0;
                length = content.ChildNodes.Count;

                foreach (XmlNode n in content.ChildNodes)
                {
                    // System.Windows.Forms.MessageBox.Show(n.OuterXml);

                    
                    counter++;
                    if (counter == length)
                        response.Append(n.OuterXml);

                }
            }
            catch (Exception e)
            {
                response.Append("Error " + e.Message);
            }
            return response.ToString();
        }

        public static string ConvertToWPMLFromTextIdx(string wordprocessingML, int idx)
        {
            //StringBuilder tmpResponse = new StringBuilder();
            StringBuilder newResponse = new StringBuilder();
            //StringBuilder response = new StringBuilder();
            try
            {
                XmlDocument document = new XmlDocument();
                document.LoadXml(wordprocessingML);
                XmlNode content = document.SelectSingleNode(bodyXPath, NamespaceManager);

                int nodeCount = content.ChildNodes.Count;
                //System.Windows.Forms.MessageBox.Show("NODE COUNT" + nodeCount);
                //removes w:sectPr
                if (content.ChildNodes.Count > 0)
                {
                    content.RemoveChild(content.LastChild);
                }

                int length = 0;
                int counter = 0;
                length = content.ChildNodes.Count;

                foreach (XmlNode n in content.ChildNodes)
                {
                    // System.Windows.Forms.MessageBox.Show(n.OuterXml);
                    if(counter == idx)
                      newResponse.Append(n.OuterXml);

                    counter++;
                   // if (counter < length)
                     //   newResponse.Append("U+016000");

                }


                //       System.Windows.Forms.MessageBox.Show("FULL STRING: " + newResponse.ToString());

                //        using (XmlTextWriter writer = new XmlTextWriter(
                //            new StringWriter(response)))
                //       {
                //            writer.Formatting = Formatting.Indented;
                //            content.WriteContentTo(writer);
                //        }
            }
            catch (Exception e)
            {
                newResponse.Append("Error " + e.Message);
            }
            return newResponse.ToString();
        }
/*
        public static string ConvertToWPMLDelimitedFromText(string wordprocessingML)
        {
            //StringBuilder tmpResponse = new StringBuilder();
            StringBuilder newResponse = new StringBuilder();
            //StringBuilder response = new StringBuilder();
            try
            {
                XmlDocument document = new XmlDocument();
                document.LoadXml(wordprocessingML);
                XmlNode content = document.SelectSingleNode(bodyXPath, NamespaceManager);
                //removes w:sectPr
                if (content.ChildNodes.Count > 0)
                {
                    content.RemoveChild(content.LastChild);
                }

                int length = 0;
                int counter = 0;
                length = content.ChildNodes.Count;

                foreach (XmlNode n in content.ChildNodes)
                {
                   // System.Windows.Forms.MessageBox.Show(n.OuterXml);
                  
                    newResponse.Append(n.OuterXml);
                    counter++;
                    if (counter < length)
                        newResponse.Append("U+016000");

                }
           

         //       System.Windows.Forms.MessageBox.Show("FULL STRING: " + newResponse.ToString());

        //        using (XmlTextWriter writer = new XmlTextWriter(
        //            new StringWriter(response)))
         //       {
        //            writer.Formatting = Formatting.Indented;
        //            content.WriteContentTo(writer);
        //        }
         }
           catch (Exception e)
            {
               newResponse.Append("Error " + e.Message);
           }
            return newResponse.ToString();
        }

*/
        //USING THIS TO INSERT FINAL 2
        //UPDATED THIS, USED FOR RETAINING SOURCE FORMATTING
        //USING FOR INSERTING WITH STYLES
        internal static string ConvertToWPMLBlock(string stylesXml, string blockXml, int idxStart, int idxEnd)
        {
            StringBuilder builder = new StringBuilder();
            XmlDocument document = new XmlDocument();
            XmlDocument paragraph = new XmlDocument();
            bool updStyles = false;

            if (!stylesXml.Equals(""))
                updStyles = true;

            try
            {
                //get entire document
                Globals.ThisAddIn.Application.ActiveDocument.Content.Select();
                string docxml = Globals.ThisAddIn.Application.Selection.WordOpenXML;
                document.LoadXml(docxml);

                //now have to reset range, currently whole document selected.
                //Works to insert after current currently selected paragraph
                //now try to insert in middle of sentence. (in middle of paragraph)
       
                object start = idxStart;
                object end = idxEnd;
                Word.Range rng = Globals.ThisAddIn.Application.ActiveDocument.Range(ref start, ref end);
                rng.Select();
             
                XmlNode body = document.SelectSingleNode(bodyXPath, NamespaceManager);
                body.InnerXml = blockXml;

                if (updStyles)
                {
                    XmlNode style = document.SelectSingleNode(stylesXPath, NamespaceManager);
                    style.InnerXml = stylesXml;
                }

                using (StringWriter writer = new StringWriter(builder))
                {
                    document.Save(writer);
                }

              
            }
            catch (Exception e)
            {
                string errMsg = e.Message;
                //System.Windows.Forms.MessageBox.Show("Error in ConvertToWordprocessingMLPara" + e.Message + " " + e.StackTrace);
            }
            return builder.ToString();
        }

        //USING THIS TO INSERT FINAL 1
        internal static string ConvertToWPMLBlock(string stylesXml, string paraXml, int paraIdx, int sentIdx, int charidx)
        {
            StringBuilder builder = new StringBuilder();
            XmlDocument document = new XmlDocument();
            XmlDocument paragraph = new XmlDocument();
            bool updStyles = false;

            if (!stylesXml.Equals(""))
                updStyles = true;

            try
            {
                
                Globals.ThisAddIn.Application.ActiveDocument.Content.Select();
                string docxml = Globals.ThisAddIn.Application.Selection.WordOpenXML;
                document.LoadXml(docxml);

                //now have to reset range, currently whole document selected.
                //Works to insert after current currently selected paragraph
                //now try to insert in middle of sentence. (in middle of paragraph).
                int tmpstart = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[1].Range.Start;
                int tmpend = Globals.ThisAddIn.Application.ActiveDocument.Paragraphs[paraIdx].Range.End;

                int diff = tmpend - tmpstart;
                int charstartrange = 0;

                if (charidx + 1 == tmpend)
                {
                    charstartrange = tmpend;
                }
                else if (charidx < tmpend)
                {
                    charstartrange = charidx - tmpstart;
                    //System.Windows.Forms.MessageBox.Show("Charrange(moving) : " + charstartrange + "tmpstart : " + tmpstart + "tmpend " + tmpend);
                }
                else
                {
                    charstartrange = diff;
                }


                object start = tmpstart;
                object end = tmpend;
                Word.Range rng = Globals.ThisAddIn.Application.ActiveDocument.Range(ref start, ref end);
                object unit = Word.WdUnits.wdCharacter;
                object count = charstartrange;
                rng.MoveStart(ref unit, ref count);
                rng.End = charstartrange;
                rng.Select();

                XmlNode body = document.SelectSingleNode(bodyXPath, NamespaceManager);
                body.InnerXml = paraXml;

                if (updStyles)
                {
                    XmlNode styles = document.SelectSingleNode(stylesXPath, NamespaceManager);
                    styles.InnerXml = stylesXml;
                }

                using (StringWriter writer = new StringWriter(builder))
                {
                    document.Save(writer);
                }

            }
            catch (Exception e)
            {
                string errMsg = e.Message;
                //System.Windows.Forms.MessageBox.Show("Error in ConvertToWordprocessingMLPara" + e.Message + " " + e.StackTrace);
            }
            return builder.ToString();
        }
        //END USING THIS TO INSERT WITH STYLES

  
    }

}
