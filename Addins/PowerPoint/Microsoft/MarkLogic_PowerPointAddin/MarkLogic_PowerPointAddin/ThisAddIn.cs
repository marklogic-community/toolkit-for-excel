using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace MarkLogic_PowerPointAddin
{
    public partial class ThisAddIn
    {
        public Microsoft.Office.Tools.CustomTaskPane myPane;
        private UserControl1 uc;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            UserControl1 uc = new UserControl1();
            myPane = this.CustomTaskPanes.Add(uc, "Mark Logic Excelerator");
            myPane.Width = 450;
            myPane.Visible = true;

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
