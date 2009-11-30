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
 * ThisAddIn.cs - The host for the Custom Task Pane and management of the pane within multiple Word instances
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Tools = Microsoft.Office.Tools;

using Extensibility;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Resources;
using System.Drawing;
using System.Windows.Forms;





namespace MarkLogic_WordAddin
{
    public partial class ThisAddIn
    {
        private AddinConfiguration ac = AddinConfiguration.GetInstance();
        private Microsoft.Office.Tools.CustomTaskPane ctpML = null;
        private Microsoft.Office.Tools.CustomTaskPane ctp = null;
        public bool mlPaneDisplayed = false;
        Document mdoc;
      

        public void AddAllTaskPanes()
        {
            
            if (Globals.ThisAddIn.Application.Documents.Count >= 0)
            {
                if (this.Application.ShowWindowsInTaskbar == true)
                {
                    foreach (Document _doc in this.Application.Documents)
                    {
                        AddTaskPane(_doc);
                    }
                }
                else
                {
                    if (!mlPaneDisplayed)
                    {
                        AddTaskPane(this.Application.ActiveDocument);
                    }
                }
               mlPaneDisplayed  = true;
            }
        }

        public void AddTaskPane(Document doc)
        {
            ctpML = this.CustomTaskPanes.Add(new UserControl1(), ac.getCTPTitleLabel(), doc.ActiveWindow);
            ctpML.Visible = true;
            ctpML.Width = 400;
          
          
        }

        public void RemoveAllTaskPanes()
        {
            if (Globals.ThisAddIn.Application.Documents.Count > 0)
            {
                for (int i = this.CustomTaskPanes.Count; i > 0; i--)
                {
                    ctp = this.CustomTaskPanes[i - 1];
                    if (ctp.Title == ac.getCTPTitleLabel())
                    {
                        this.CustomTaskPanes.RemoveAt(i - 1);
                    }
                }
                mlPaneDisplayed = false;
            }
        }

        private void RemoveOrphanedTaskPanes()
        {
            for (int i = this.CustomTaskPanes.Count; i > 0; i--)
            {
                ctp = this.CustomTaskPanes[i - 1];
                if (ctp.Window == null)
                {
                    this.CustomTaskPanes.Remove(ctp);
                }
            }
        }

        //public delegate void EventNameEventHandler(object sender, EventNameEventArgs e);

        private void Application_DocumentBeforeClose(Document doc, ref bool cancel)
        {
           /* if (mdoc.Name == doc.Name)
            {
               // assumes you added all of the controls in the document
               // foreach (Word.ContentControl control in doc.ContentControls)
               // {
               //      delete control - but do not delete contents
               //   control.Delete(false);
               // }
               
            }*/
            cancel = false;
        }

        private void Application_DocumentOpen(Document Doc)
        {
           //MessageBox.Show("IN OPEN");
            RemoveOrphanedTaskPanes();
           // mdoc = null;
           // mdoc = Doc;
           // mdoc.ContentControlOnEnter += new DocumentEvents2_ContentControlOnEnterEventHandler(this.ThisDocument_ContentControlOnEnter);
           // mdoc.ContentControlOnExit += new DocumentEvents2_ContentControlOnExitEventHandler(this.ThisDocument_ContentControlOnExit);

           
            if (mlPaneDisplayed && this.Application.ShowWindowsInTaskbar)
            {
                AddTaskPane(Doc);
            }
        }

        private void Application_NewDocument(Document Doc)
        {
           // MessageBox.Show("IN NEW");
           // mdoc = null;
           // mdoc = this.Application.ActiveDocument;
           // mdoc.ContentControlOnEnter += new DocumentEvents2_ContentControlOnEnterEventHandler(this.ThisDocument_ContentControlOnEnter);
           // mdoc.ContentControlOnExit += new DocumentEvents2_ContentControlOnExitEventHandler(this.ThisDocument_ContentControlOnExit);

            if (mlPaneDisplayed && this.Application.ShowWindowsInTaskbar)
            {
                AddTaskPane(Doc);
            }
        }

        private void Application_DocumentChange()
        {
            //MessageBox.Show("IN CHANGE");

            RemoveOrphanedTaskPanes();
            mdoc = null;
            mdoc = this.Application.ActiveDocument;
            mdoc.ContentControlOnEnter += new DocumentEvents2_ContentControlOnEnterEventHandler(this.ThisDocument_ContentControlOnEnter);
            mdoc.ContentControlOnExit += new DocumentEvents2_ContentControlOnExitEventHandler(this.ThisDocument_ContentControlOnExit);

            //MessageBox.Show("Changing");
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string ribbonBtnLabel = ac.getRibbonButtonLabel();
            string ribbonGroupLabel = ac.getRibbonGroupLabel();
            string ribbonTabLabel = ac.getRibbonTabLabel();
            if(!(ribbonBtnLabel.Equals("") || ribbonBtnLabel==null)) 
               Globals.Ribbons.Ribbon1.viewTaskPaneButton.Label = ribbonBtnLabel;

            if (!(ribbonTabLabel.Equals("") || ribbonTabLabel == null)) 
               Globals.Ribbons.Ribbon1.tab2.Label = ribbonTabLabel;

            if (!(ribbonGroupLabel.Equals("") || ribbonGroupLabel == null)) 
               Globals.Ribbons.Ribbon1.group1.Label = ribbonGroupLabel;

            this.Application.DocumentOpen +=
                new Microsoft.Office.Interop.Word.
                ApplicationEvents4_DocumentOpenEventHandler(
                Application_DocumentOpen);

            ((ApplicationEvents4_Event)this.Application).NewDocument +=
                new Microsoft.Office.Interop.Word.
                ApplicationEvents4_NewDocumentEventHandler(
                Application_NewDocument);

            this.Application.DocumentChange +=
                new Microsoft.Office.Interop.Word.
                ApplicationEvents4_DocumentChangeEventHandler(
                Application_DocumentChange);

            this.Application.DocumentBeforeClose += new ApplicationEvents4_DocumentBeforeCloseEventHandler(Application_DocumentBeforeClose);

            if (this.Application.ShowWindowsInTaskbar == true)
            {
                if (ac.getPaneEnabled())
                {

                    Globals.Ribbons.Ribbon1.viewTaskPaneButton.Checked = true;
                    AddAllTaskPanes();
                }
            }

            
          //  Globals.ThisAddIn.Application.ActiveDocument.ContentControlOnEnter += new DocumentEvents2_ContentControlOnEnterEventHandler(this.ThisDocument_ContentControlOnEnter);
          //  Globals.ThisAddIn.Application.ActiveDocument.ContentControlOnExit += new DocumentEvents2_ContentControlOnExitEventHandler(this.ThisDocument_ContentControlOnExit);

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            RemoveAllTaskPanes();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        
        }

        
    private void ThisDocument_ContentControlOnEnter(ContentControl contentControl)
        {
            UserControl1 uc = (UserControl1)this.ctpML.Control;
            string parentTag="";
            string parentID = "";

            try
            {
                ContentControl parent = contentControl.ParentContentControl;
                parentTag = parent.Tag;
                parentID = parent.ID;

            }
            catch (Exception e)
            {
                //do nothing, not parent
                string donothing_removewarning = e.Message;
                //MessageBox.Show("No Parent");
            }
            uc.contentControlOnEnter( contentControl.ID, contentControl.Tag, contentControl.Title, contentControl.Type.ToString(), parentTag, parentID);
            //webBrowser1.Document.InvokeScript("testOnEnter",  new String[] { "called from client code" });
            //MessageBox.Show(String.Format( "ContentControl of type {0} with ID {1} and Tag {2} entered.",contentControl.Type, contentControl.ID, contentControl.Tag));

        }

        private void ThisDocument_ContentControlOnExit(ContentControl contentControl, ref bool cancel)
        {
           // this.CustomTaskPanes.Add(new UserControl1(), ac.getCTPTitleLabel(), doc.ActiveWindow);
            UserControl1 uc = (UserControl1)this.ctpML.Control;
            string parentTag = "";
            string parentID = "";

            try
            {
                ContentControl parent = contentControl.ParentContentControl;
                parentTag = parent.Tag;
                parentID = parent.ID;

            }
            catch (Exception e)
            {
                //do nothing, not parent
                string donothing_removewarning = e.Message;
                //MessageBox.Show("No Parent");
            }

            uc.contentControlOnExit(contentControl.ID, contentControl.Tag, contentControl.Title, contentControl.Type.ToString(), parentTag, parentID);
            //uc.webBrowser1.Document.InvokeScript("testOnExit", new String[] { "called from client code 2" });
            //MessageBox.Show("Exiting");
           // MessageBox.Show(String.Format(

             // "ContentControl of type {0} with ID {1} exited.",

              //contentControl.Type, contentControl.ID));

        }
    
        //test
        /*
               public void OnConnection(object application,
                                 Extensibility.ext_ConnectMode connectMode,
                                 object addInInst, ref System.Array custom)
                {
                   // addInInst = this;
                    Microsoft.VisualBasic.Interaction.CallByName(addInInst, "Object", Microsoft.VisualBasic.CallType.Let, this);
                }
         */

        #endregion
    }
}
