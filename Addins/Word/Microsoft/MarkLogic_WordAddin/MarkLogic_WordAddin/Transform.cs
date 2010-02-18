/*Copyright 2008-2010 Mark Logic Corporation

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
 * Transform.cs - Used to insert/extract WordprocessingML to/from the active document.
*/
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

        public static string GetActiveDocumentXml(string wordprocessingML)
        {
            StringBuilder response = new StringBuilder();
            try
            {
                XmlDocument document = new XmlDocument();
                document.LoadXml(wordprocessingML);
                XmlNode content = document.SelectSingleNode(documentXPath, NamespaceManager);
               
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

        public static string getRows(XmlNodeList nodes, string delimiter)
        {
            string response = "";
            string delim = "";
            if (delimiter.Equals(""))
                delim = "\t";
            else delim = delimiter;

            foreach (XmlNode n in nodes)
            {
                 switch (n.NodeType)
                {

                    case XmlNodeType.Element:
                        if (n.Name.Equals("w:tr"))
                        {
                            XmlNodeList cells = n.SelectNodes("w:tc", NamespaceManager);
                            foreach (XmlNode cell in cells)
                                response += cell.InnerText + delim;

                        }

                        break;
                }
            }

            return response + "\n"; ;
        }

        public static string ExtractTextValuesFromXML(string wordprocessingML, int idx, string delimiter)
        {
            StringBuilder response = new StringBuilder();

            try
            {

                XmlDocument document = new XmlDocument();

                document.LoadXml(wordprocessingML);
                XmlNode content = document.SelectSingleNode(bodyXPath, NamespaceManager);

                int nodeCount = content.ChildNodes.Count;
                //removes w:sectPr
                if (content.ChildNodes.Count > 0)
                {
                    content.RemoveChild(content.LastChild);
                }

                int length = content.ChildNodes.Count;
                int counter = 0;
                foreach (XmlNode n in content.ChildNodes)
                {
                if (counter == idx)
                {
                        switch (n.NodeType)
                        {

                            case XmlNodeType.Element:
                                if (n.Name.Equals("w:tbl"))
                                {
                                    response.Append(getRows(n.ChildNodes, delimiter));
                                }
                                else if (n.Name.Equals("w:p"))
                                {
                                    response.Append(n.InnerText+"\n");
                                }
                                else if (n.Name.Equals("w:sdt"))
                                {   
                                    XmlNodeList xl = n.SelectNodes("descendant::w:sdtContent[1]", NamespaceManager);
                                    XmlNodeList tmp = xl;
                                Found:
                                    foreach (XmlNode xn in tmp)
                                    {
                                        foreach (XmlNode child in xn.ChildNodes)
                                        {
                                            if (child.Name.Equals("w:sdt"))
                                            {
                                                tmp = child.SelectNodes("descendant::w:sdtContent[1]", NamespaceManager);
                                                goto Found;
                                            }
                                            else if (child.Name.Equals("w:tbl"))
                                            { 
                                                response.Append(getRows(child.ChildNodes, delimiter));
                                            }
                                            else if (child.Name.Equals("w:p"))
                                            {
                                                response.Append(child.InnerText + "\n");
                                            }
                                            else if (child.Name.Equals("w:customXml"))
                                            {
                                                response.Append(child.InnerText + "\n");
                                            }
   
                                        }
                                    }
                                
                                }
                                else if (n.Name.Equals("w:customXml"))
                                {
                                    response.Append(n.InnerText);
                                }
                                break;
                        }
                   }
                   counter++;
                }

            }
            catch (Exception e)
            {
                response.Append("Error " + e.Message);
            }

            return response.ToString();

        }

        //get selected xml as text
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

        //get selected xml as text by index
        public static string ConvertToWPMLFromTextIdx(string wordprocessingML, int idx)
        {
            StringBuilder newResponse = new StringBuilder();

            try
            {
                XmlDocument document = new XmlDocument();
                document.LoadXml(wordprocessingML);
                XmlNode content = document.SelectSingleNode(bodyXPath, NamespaceManager);

                int nodeCount = content.ChildNodes.Count;

                if (content.ChildNodes.Count > 0)
                {
                    content.RemoveChild(content.LastChild);
                }

                int length = 0;
                int counter = 0;
                length = content.ChildNodes.Count;

                foreach (XmlNode n in content.ChildNodes)
                {
                    if(counter == idx)
                      newResponse.Append(n.OuterXml);

                    counter++;

                }

            }
            catch (Exception e)
            {
                newResponse.Append("Error " + e.Message);
            }

            return newResponse.ToString();
        }

        //used to insert block level element into body
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
            }
            return builder.ToString();
        }

        //used to set packagexml
        internal static string setPackageXML(string documentXml)
        {
            StringBuilder builder = new StringBuilder();
            XmlDocument document = new XmlDocument();

            try
            {
                Globals.ThisAddIn.Application.ActiveDocument.Content.Select();


                //write to file here:
                /*  System.Windows.Forms.MessageBox.Show(docx);
                  TextWriter tw = new StreamWriter(@"C:\origdocx.xml");
                  tw.WriteLine(docx);
                  tw.Close();

                  System.Windows.Forms.MessageBox.Show(documentXml);
                  TextWriter tw2 = new StreamWriter(@"C:\sdtxml.xml");
                  tw2.WriteLine(documentXml);
                  tw2.Close();
                 * */

                document.LoadXml(documentXml);

               // XmlNode doc = document.SelectSingleNode(documentXPath, NamespaceManager);
                //doc.InnerXml = documentXml;

                using (StringWriter writer = new StringWriter(builder))
                {
                    document.Save(writer);
                }

            }
            catch (Exception e)
            {
                string errMsg = e.Message;
            }

            return builder.ToString();
        }

        //used to set document.xml
        internal static string ConvertToWPML(string documentXml)
        {
            StringBuilder builder = new StringBuilder();
            XmlDocument document = new XmlDocument();

            try
            {
                Globals.ThisAddIn.Application.ActiveDocument.Content.Select();

                string docx = Globals.ThisAddIn.Application.Selection.WordOpenXML;

                //write to file here:
              /*  System.Windows.Forms.MessageBox.Show(docx);
                TextWriter tw = new StreamWriter(@"C:\origdocx.xml");
                tw.WriteLine(docx);
                tw.Close();

                System.Windows.Forms.MessageBox.Show(documentXml);
                TextWriter tw2 = new StreamWriter(@"C:\sdtxml.xml");
                tw2.WriteLine(documentXml);
                tw2.Close();
               * */

                document.LoadXml(docx);

                XmlNode doc = document.SelectSingleNode(documentXPath, NamespaceManager);
                doc.InnerXml = documentXml;

                using (StringWriter writer = new StringWriter(builder))
                {
                    document.Save(writer);
                }

            }
            catch (Exception e)
            {
                string errMsg = e.Message;
            }

            return builder.ToString();
        }

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
            }
            return builder.ToString();
        }
  
    }

}
