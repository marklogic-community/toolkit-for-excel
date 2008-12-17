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
        private Microsoft.Office.Tools.CustomTaskPane ctpML = null;
        private Microsoft.Office.Tools.CustomTaskPane ctp = null;
        public bool mlPaneDisplayed = false;
      

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
