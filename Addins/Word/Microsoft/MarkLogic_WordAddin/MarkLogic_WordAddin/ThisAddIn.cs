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
        CTPManager<IonControl1> _manager = null;
        public CTPManager<IonControl1> CTPManager
        {
            get
            {
                if (_manager == null)
                {
                    _manager = new CTPManager<IonControl1>(
                        new TaskPaneFactory());

                }
                return _manager;
            }
        }


        internal void ClearTaskPanes()
        {
            foreach (IonControl1 taskPane in CTPManager.GetTaskPanes())
            {
                taskPane.Clear();
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Globals.ThisAddIn.CTPManager.ToggleTaskPane(Globals.ThisAddIn.Application.ActiveWindow);
            
            //This is setting pane open on startupe
            Globals.Ribbons.Ribbon1.viewTaskPaneButton.Checked = true;

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        
        class TaskPaneFactory
           : CTPManager<IonControl1>.ITaskPaneFactory
        {
            public string CreateTitle(IonControl1 taskPane)
            {
                return "Mark Logic Authoring Kit";
            }

            public IonControl1 CreateNewTaskPane(
                Document document, Window window)
            {
                IonControl1 pane = new IonControl1();
                pane.Document = document;
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
