using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tools = Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace MarkLogic_ExcelAddin
{
    public class CTPManager<TTaskPane> where TTaskPane : UserControl1
    {
        Dictionary<Excel.Window, TaskPaneEntry> custTaskPanes = null;
        ITaskPaneFactory custTaskPaneFactory = null;
        List<RibbonToggleButton> managedButtons = null;

        struct TaskPaneEntry
        {
            public TTaskPane UserTaskPane;
            public Tools.CustomTaskPane MLTaskPane;
        }

        public CTPManager(ITaskPaneFactory taskPaneFactory)
        {
            custTaskPanes = new Dictionary<Excel.Window, TaskPaneEntry>();
            custTaskPaneFactory = taskPaneFactory;
            managedButtons = new List<RibbonToggleButton>();
            Globals.ThisAddIn.Application.WindowActivate += new Microsoft.Office.Interop.Excel.AppEvents_WindowActivateEventHandler(Application_WindowActivate);
            //Globals.ThisAddIn.Application.WindowActivate += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowActivateEventHandler(Application_WindowActivate);

        }

        public void ManageToggleButton(RibbonToggleButton toggleButton)
        {
            managedButtons.Add(toggleButton);
            toggleButton.Click +=
                delegate(object sender, RibbonControlEventArgs e)
                {
                    ToggleTaskPane(Globals.ThisAddIn.Application.ActiveWindow);
                };
        }

        public TTaskPane GetCurrentTaskPane()
        {
            TTaskPane pane = null;
            Excel.Window activeWindow = Globals.ThisAddIn.Application.ActiveWindow;
            if (custTaskPanes.ContainsKey(activeWindow))
            {
                pane = custTaskPanes[activeWindow].UserTaskPane;
            }
            return pane;
        }

        public bool IsTaskPaneVisible(Excel.Window window)
        {
            bool visible = false;
            if (custTaskPanes.ContainsKey(window))
            {
                visible = custTaskPanes[window].MLTaskPane.Visible;
            }
            return visible;
        }

        public TTaskPane ToggleTaskPane(Excel.Window window)
        {
            TaskPaneEntry entry;
            if (custTaskPanes.ContainsKey(window))
            {
                entry = custTaskPanes[window];
            }
            else
            {
                entry = new CTPManager<TTaskPane>.TaskPaneEntry();
              //  TTaskPane taskPane = custTaskPaneFactory.CreateNewTaskPane(
              //            window.Document, window);
                TTaskPane taskPane = custTaskPaneFactory.CreateNewTaskPane(
                   window.Application.Workbooks, window);
                entry.UserTaskPane = taskPane;
                entry.MLTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(
                    taskPane, custTaskPaneFactory.CreateTitle(taskPane), window);
                entry.MLTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                //Width of Pane
                entry.MLTaskPane.Width = 400;
                entry.MLTaskPane.VisibleChanged += new EventHandler(VstoTaskPane_VisibleChanged);
                custTaskPanes.Add(window, entry);
            }
            entry.MLTaskPane.Visible = !entry.MLTaskPane.Visible;
            RefreshToggleButtons();
            return entry.UserTaskPane;
        }

        internal IEnumerable<TTaskPane> GetTaskPanes()
        {
            List<TTaskPane> taskPanes = new List<TTaskPane>();
            foreach (TaskPaneEntry entry in custTaskPanes.Values)
            {
                taskPanes.Add(entry.UserTaskPane);
            }
            return taskPanes;
        }

        void Application_WindowActivate(Microsoft.Office.Interop.Excel.Workbook wkbk, Microsoft.Office.Interop.Excel.Window Wn)
        {
            RefreshToggleButtons();
        }

        void RefreshToggleButtons()
        {
            foreach (RibbonToggleButton button in managedButtons)
            {
                button.Checked = IsTaskPaneVisible(Globals.ThisAddIn.Application.ActiveWindow);
            }
        }

        void VstoTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            Tools.CustomTaskPane taskPane = (Tools.CustomTaskPane)sender;
            TTaskPane userTaskPane = custTaskPanes[(Excel.Window)taskPane.Window].UserTaskPane;
            RefreshToggleButtons();
            OnTaskPaneVisibilityChanged(
                new CTPManager<TTaskPane>.TaskPaneVisiblityChangedEventArgs(userTaskPane));
        }

        void OnTaskPaneVisibilityChanged(TaskPaneVisiblityChangedEventArgs e)
        {
            EventHandler<TaskPaneVisiblityChangedEventArgs> handler = TaskPaneVisibilityChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event EventHandler<TaskPaneVisiblityChangedEventArgs> TaskPaneVisibilityChanged;

        public interface ITaskPaneFactory
        {
            string CreateTitle(TTaskPane taskPane);
            TTaskPane CreateNewTaskPane(Excel.Workbooks wkbk, Excel.Window window);
        }

        public class TaskPaneVisiblityChangedEventArgs
            : EventArgs
        {
            public TTaskPane TaskPane { get; private set; }

            public TaskPaneVisiblityChangedEventArgs(TTaskPane taskPane)
            {
                TaskPane = taskPane;
            }
        }
    }


}