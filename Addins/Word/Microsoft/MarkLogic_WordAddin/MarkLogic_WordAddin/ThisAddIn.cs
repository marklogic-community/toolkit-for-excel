/*Copyright 2002-2008 Mark Logic Corporation.  All Rights Reserved*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Tools = Microsoft.Office.Tools;




namespace MarkLogic_WordAddin
{
    public partial class ThisAddIn
    {
        private AddinConfiguration ac = AddinConfiguration.GetInstance();
        //private string ctpTitle = "";

        CTPManager<UserControl1> mgr = null;
        public CTPManager<UserControl1> CTPManager
        {
            get
            {
                if (mgr == null)
                {
                    mgr = new CTPManager<UserControl1>(
                        new CustomTaskPaneFactory());
                   
                }
                return mgr;
            }
        }

        internal void ClearTaskPanes()
        {
            foreach (UserControl1 taskPane in CTPManager.GetTaskPanes())
            {
                taskPane.Clear();
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        { 
            string ribbonBtnLabel = ac.getRibbonButtonLabel();
            string ribbonGroupLabel = ac.getRibbonGroupLabel();
            string ribbonTabLabel = ac.getRibbonTabLabel();
            //Globals.Ribbons.Ribbon1.viewTaskPaneButton.Checked = true;
            //Globals.ThisAddIn.CTPManager.ManageToggleButton(Globals.Ribbons.Ribbon1.viewTaskPaneButton);
            //Globals.ThisAddIn.CTPManager.ToggleTaskPane(Globals.ThisAddIn.Application.ActiveWindow);
            //Globals.Ribbons.Ribbon1.viewTaskPaneButton.Checked = true;
            if(!(ribbonBtnLabel.Equals("") || ribbonBtnLabel==null)) 
               Globals.Ribbons.Ribbon1.viewTaskPaneButton.Label = ribbonBtnLabel;

            if (!(ribbonTabLabel.Equals("") || ribbonTabLabel == null)) 
               Globals.Ribbons.Ribbon1.tab2.Label = ribbonTabLabel;

            if (!(ribbonGroupLabel.Equals("") || ribbonGroupLabel == null)) 
               Globals.Ribbons.Ribbon1.group1.Label = ribbonGroupLabel;

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        
        class CustomTaskPaneFactory
           : CTPManager<UserControl1>.ITaskPaneFactory
        {
            private AddinConfiguration ac = AddinConfiguration.GetInstance();
            public string CreateTitle(UserControl1 taskPane)
            {
                string title = ac.getCTPTitleLabel();
                if (title.Equals("") || title == null)
                    return "Mark Logic Authoring Kit";
                else return title;
                // return "Mark Logic Authoring Kit";
            }

            public UserControl1 CreateNewTaskPane(
                Document document, Window window)
            {
                UserControl1 pane = new UserControl1();
                pane.Document = document;
               // Globals.Ribbons.Ribbon1.viewTaskPaneButton.Checked = true;
               // pane.Visible = true;
          
                return pane;
            }
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
