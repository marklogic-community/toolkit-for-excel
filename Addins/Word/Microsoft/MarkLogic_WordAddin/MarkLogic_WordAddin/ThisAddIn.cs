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
        private Document mdoc;
        private bool debug = false;
      
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
            ctpML = this.CustomTaskPanes.Add(new UserControl1(doc), ac.getCTPTitleLabel(), doc.ActiveWindow);
            ctpML.Visible = true;
            ctpML.Width = 400;
        }

        public void RemoveAllTaskPanes()
        {
            if(debug)
                MessageBox.Show("RemoveAllTaskPanes()");

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
            //MessageBox.Show("In this remove orphaned task panes");

            for (int i = this.CustomTaskPanes.Count; i > 0; i--)
            {
                ctp = this.CustomTaskPanes[i - 1];
                if (ctp.Window == null)
                {
                    this.CustomTaskPanes.Remove(ctp);
                }
            }
        }

        private void Application_DocumentBeforeClose(Document doc, ref bool cancel)
        {
            if(debug)
                MessageBox.Show("begin Application_DocumentBeforeClose");

            mdoc = doc;
            UserControl1 uc = (UserControl1)this.ctpML.Control;

            mdoc.ContentControlOnEnter -=  new DocumentEvents2_ContentControlOnEnterEventHandler(uc.ThisDocument_ContentControlOnEnter);
            mdoc.ContentControlOnExit -=  new DocumentEvents2_ContentControlOnExitEventHandler(uc.ThisDocument_ContentControlOnExit);
            mdoc.ContentControlAfterAdd -=  new DocumentEvents2_ContentControlAfterAddEventHandler(uc.ThisDocument_ContentControlAfterAdd);
            mdoc.ContentControlBeforeDelete -= new DocumentEvents2_ContentControlBeforeDeleteEventHandler(uc.ThisDocument_ContentControlBeforeDelete);
            mdoc.ContentControlBeforeContentUpdate -= new DocumentEvents2_ContentControlBeforeContentUpdateEventHandler(uc.ThisDocument_ContentControlBeforeContentUpdate);
            mdoc.ContentControlBeforeStoreUpdate -= new DocumentEvents2_ContentControlBeforeStoreUpdateEventHandler(uc.ThisDocument_ContentControlBeforeStoreUpdate);

            if(debug)
                MessageBox.Show("end Application_DocumentBeforeClose");

            cancel = false;
        }

        private void Application_DocumentOpen(Document Doc)
        {
            if(debug)
                MessageBox.Show("Application_DocumentOpen");

            RemoveOrphanedTaskPanes();

            if (mlPaneDisplayed && this.Application.ShowWindowsInTaskbar)
            {
                AddTaskPane(Doc);
            }
        }

        private void Application_NewDocument(Document Doc)
        {
            if(debug)
                MessageBox.Show("Application_NewDocument");

            if (mlPaneDisplayed && this.Application.ShowWindowsInTaskbar)
            {
                AddTaskPane(Doc);
            }
        }

        private void Application_DocumentChange()
        {
            if(debug)
                MessageBox.Show("Application_DocumentChange");

            RemoveOrphanedTaskPanes();

        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if(debug)
                 MessageBox.Show("ThisAddin_Startup");

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

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if(debug)
                 MessageBox.Show("ThisAddIn_Shutdown");

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

        #endregion
    }
}
