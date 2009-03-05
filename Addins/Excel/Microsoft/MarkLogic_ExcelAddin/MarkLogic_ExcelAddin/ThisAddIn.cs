using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;


using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;


namespace MarkLogic_ExcelAddin
{
    public partial class ThisAddIn
    {
        public Microsoft.Office.Tools.CustomTaskPane myPane;
        private UserControl1 uc;
        private AddinConfiguration ac = AddinConfiguration.GetInstance();
       
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            string ribbonBtnLabel = ac.getRibbonButtonLabel();
            string ribbonGroupLabel = ac.getRibbonGroupLabel();
            string ribbonTabLabel = ac.getRibbonTabLabel();
            if (!(ribbonBtnLabel.Equals("") || ribbonBtnLabel == null))
                Globals.Ribbons.Ribbon1.viewTaskPaneButton.Label = ribbonBtnLabel;

            if (!(ribbonTabLabel.Equals("") || ribbonTabLabel == null))
                Globals.Ribbons.Ribbon1.tab2.Label = ribbonTabLabel;

            if (!(ribbonGroupLabel.Equals("") || ribbonGroupLabel == null))
                Globals.Ribbons.Ribbon1.group2.Label = ribbonGroupLabel;

            UserControl1 uc = new UserControl1();
            myPane = this.CustomTaskPanes.Add(uc, ac.getCTPTitleLabel()); //"Mark Logic Excelerator");
            myPane.Width = 325;
            myPane.Visible = ac.getPaneEnabled();

            myPane.VisibleChanged += new EventHandler(taskPaneValue_VisibleChanged);

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
         
        }

        private void taskPaneValue_VisibleChanged(object sender, System.EventArgs e)
        {
            Globals.Ribbons.Ribbon1.viewTaskPaneButton.Checked = myPane.Visible;
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
