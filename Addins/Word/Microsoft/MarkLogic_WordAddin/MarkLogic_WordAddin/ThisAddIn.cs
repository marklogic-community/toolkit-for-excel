/*Copyright 2002-2008 Mark Logic Corporation.  All Rights Reserved*/
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
        //ctpCalendar
        private Microsoft.Office.Tools.CustomTaskPane ctpML = null;
        private Microsoft.Office.Tools.CustomTaskPane ctp = null;
        //private Window ctpWindow = null;
        public bool mlPaneDisplayed = false;
      

        public void AddAllTaskPanes()
        {
            
            if (Globals.ThisAddIn.Application.Documents.Count >= 0)
            {
                // If Show all windows in the Taskbar is selected then 
                // each open document has its own window.  
                // If Show all windows in the Taskbar is not selected  
                // then Word displays each open document in the same window.  
                if (this.Application.ShowWindowsInTaskbar == true)
                {
                    // Loop through each open document window
                    foreach (Document _doc in this.Application.Documents)
                    {
                        // Pass this document as a parameter to AddCustomTaskPane
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
            // Create a new custom task pane and add it to the 
            // collection of custom task panes belonging to this add-in
            // The first two arguments of the Add method specify a control to add
            // to the custom task pane and the title to display on the task pane. 
            // The third argument, which is optional, specifies the 
            // parent window for the custom task pane. 
            ctpML = this.CustomTaskPanes.Add(new UserControl1(), ac.getCTPTitleLabel(), doc.ActiveWindow);
            ctpML.Visible = true;
            ctpML.Width = 400;
        }

        public void RemoveAllTaskPanes()
        {
            // First check if there are any open documents.
            if (Globals.ThisAddIn.Application.Documents.Count > 0)
            {
                // Loop through each custom task pane belonging to the add-in
                for (int i = this.CustomTaskPanes.Count; i > 0; i--)
                {
                    ctp = this.CustomTaskPanes[i - 1];
                    if (ctp.Title == ac.getCTPTitleLabel())
                    {
                        // If this is a ml task pane, remove it
                        this.CustomTaskPanes.RemoveAt(i - 1);
                    }
                }
                mlPaneDisplayed = false;
            }
        }

        private void RemoveOrphanedTaskPanes()
        {
            // Loop through each custom task pane belonging to the add-in
            for (int i = this.CustomTaskPanes.Count; i > 0; i--)
            {
                ctp = this.CustomTaskPanes[i - 1];
                // If this task pane has no associated window, remove it
                if (ctp.Window == null)
                {
                    this.CustomTaskPanes.Remove(ctp);
                }
            }
        }

        private void Application_DocumentOpen(Document Doc)
        {
            RemoveOrphanedTaskPanes();
            if (mlPaneDisplayed && this.Application.ShowWindowsInTaskbar)
            {
                AddTaskPane(Doc);
            }
        }

        private void Application_NewDocument(Document Doc)
        {
            if (mlPaneDisplayed && this.Application.ShowWindowsInTaskbar)
            {

                AddTaskPane(Doc);
            }
        }

        private void Application_DocumentChange()
        {
            RemoveOrphanedTaskPanes();
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

           //If we add this, then have to set 
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

        //added for testing
        /*
               public void OnConnection(object application,
                                 Extensibility.ext_ConnectMode connectMode,
                                 object addInInst, ref System.Array custom)
                {
                   // addInInst = this;
                    System.Windows.Forms.MessageBox.Show("ADDIN");
                    Microsoft.VisualBasic.Interaction.CallByName(addInInst, "Object", Microsoft.VisualBasic.CallType.Let, this);

                }
         */

        #endregion
    }
}
