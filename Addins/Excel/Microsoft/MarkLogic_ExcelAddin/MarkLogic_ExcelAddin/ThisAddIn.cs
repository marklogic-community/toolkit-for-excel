/*Copyright 2009-2010 Mark Logic Corporation

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
 * ThisAddIn.cs - The host for the Custom Task Pane and management of the pane within Excel
*/
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
        private AddinConfiguration ac = AddinConfiguration.GetInstance();
        public bool mlPaneDisplayed = false;
        private UserControl1 taskPaneControl1;
        private Microsoft.Office.Tools.CustomTaskPane ctpML;
  
         
         // public delegate void ChartEvents_SelectEventHandler(int elementId, int arg1, int arg2);
         // public event Excel.ChartEvents_SelectEventHandler chartSelected;
        


   

        private void ctpML_VisibleChanged(object sender, System.EventArgs e)
        {
            
            Globals.Ribbons.Ribbon1.viewTaskPaneButton.Checked =
                ctpML.Visible;
        }

        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get
            {
                return ctpML;
            }
        }

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


            taskPaneControl1 = new UserControl1();
            ctpML = this.CustomTaskPanes.Add(
                 taskPaneControl1, ac.getCTPTitleLabel());
            ctpML.VisibleChanged +=
                new EventHandler(ctpML_VisibleChanged);
            ctpML.Width = 400;
            ctpML.Visible = ac.getPaneEnabled();
            //changing follwoing to WB level
            //Excel.Application app = Globals.ThisAddIn.Application;
            //app.EnableEvents = true;
            //app.SheetActivate+= new Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler(app_SheetActivate);
        /*THESE
         * Excel.Application app = Globals.ThisAddIn.Application;
            app.SheetActivate+=new Microsoft.Office.Interop.Excel.AppEvents_SheetActivateEventHandler(app_SheetActivate);
            app.WorkbookOpen += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookOpenEventHandler(app_WorkbookOpen
         */

            //wb.SheetActivate+=new Microsoft.Office.Interop.Excel.WorkbookEvents_SheetActivateEventHandler(wb_SheetActivate);
            //Excel.Sheets charts = app.Charts;
            //Excel.ChartObjects chartObjects = (Excel.ChartObjects)(Sheet.ChartObjects(Type.Missing));

           /* foreach (Excel.Worksheet c in charts)
            {
                try{
                Excel.ChartObjects chartObjects = (Excel.ChartObjects)(c.ChartObjects(Type.Missing));
                MessageBox.Show(chartObjects.Count+"");
                }catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            */
           /* Excel.Sheets charts = app.Charts;
            MessageBox.Show("" + charts.Count);
            int chartCnt = charts.Count;
            int idx = 1;
            foreach (Excel.Worksheet ws in charts)
            {
                Microsoft.Office.Tools.Excel.Chart c = (Microsoft.Office.Tools.Excel.Chart)ws.ChartObjects(idx);
                MessageBox.Show(c.Name);
                c.SelectEvent+=new Microsoft.Office.Interop.Excel.ChartEvents_SelectEventHandler(chart_SelectEvent);
                c.ActivateEvent += new Microsoft.Office.Interop.Excel.ChartEvents_ActivateEventHandler(chart_Activate);
                idx++;

            }
            
         */
       /*     Excel.Sheets charts = Globals.ThisAddIn.Application.Charts;
            int chartCount = charts.Count;
            int idx=1;
            foreach (Excel.Worksheet s in charts)
            {
                Excel.Chart chart = s.ChartObjects(idx);
                chart.Activate()+= new Microsoft.Office.Interop.Excel.ChartEvents_ActivateEventHandler(chart_Activate);
                idx++;
                }
                //c.Activate += new Microsoft.Office.Interop.Excel.ChartEvents_ActivateEventHandler(chart_Activate);
            }
            
            */

        }

        // another way (besides using is)to check what sheet is
  //      foreach (object obj in doc.Sheets)//
//{
// Microsoft.Office.Interop.Excel.Worksheet sheet = obj as Microsoft.Office.Interop.Excel.Worksheet;
// if (sheet != null)
  // ...
//}



       

        //consider 
        //new, newchart, newsheet events



        
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            
           
        
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
