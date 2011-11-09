/*Copyright 2009-2011 Mark Logic Corporation

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
 * ThisAddIn.cs - The host for the Custom Task Pane and management of the pane within PowerPoint
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System.Runtime.InteropServices;

namespace MarkLogic_PowerPointAddin
{
    public partial class ThisAddIn
    {
        private AddinConfiguration ac = AddinConfiguration.GetInstance();
        public bool mlPaneDisplayed = false;
        //private UserControl1 uc;
        public Microsoft.Office.Tools.CustomTaskPane myPane;


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
            myPane = this.CustomTaskPanes.Add(uc, ac.getCTPTitleLabel());
            myPane.Width = 450;
            myPane.Visible = ac.getPaneEnabled();

            myPane.VisibleChanged += new EventHandler(taskPaneValue_VisibleChanged);
        }
       

        public void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // MessageBox.Show("Quitting");

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
