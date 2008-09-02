using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tools = Microsoft.Office.Tools;
using MSWord = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace MarkLogic_WordAddin
{
    public class CTPManager<TTaskPane> where TTaskPane : UserControl1
    {
        Dictionary<MSWord.Window, TaskPaneEntry> _taskPanes = null;
        ITaskPaneFactory _taskPaneFactory = null;
        List<RibbonToggleButton> _managedButtons = null;

        public CTPManager(ITaskPaneFactory taskPaneFactory)
        {
            _taskPanes = new Dictionary<MSWord.Window, TaskPaneEntry>();
            _taskPaneFactory = taskPaneFactory;
            _managedButtons = new List<RibbonToggleButton>();
            Globals.ThisAddIn.Application.WindowActivate += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowActivateEventHandler(Application_WindowActivate);
        }

        public void ManageToggleButton(RibbonToggleButton toggleButton)
        {
            _managedButtons.Add(toggleButton);
            toggleButton.Click +=
                delegate(object sender, RibbonControlEventArgs e)
                {
                    ToggleTaskPane(Globals.ThisAddIn.Application.ActiveWindow);
                };
        }

        public TTaskPane GetCurrentTaskPane()
        {
            TTaskPane pane = null;
            MSWord.Window activeWindow = Globals.ThisAddIn.Application.ActiveWindow;
            if (_taskPanes.ContainsKey(activeWindow))
            {
                pane = _taskPanes[activeWindow].UserTaskPane;
            }
            return pane;
        }

        public bool IsTaskPaneVisible(MSWord.Window window)
        {
            bool visible = false;
            if (_taskPanes.ContainsKey(window))
            {
                visible = _taskPanes[window].MLTaskPane.Visible;
            }
            return visible;
        }

        public TTaskPane ToggleTaskPane(MSWord.Window window)
        {
            TaskPaneEntry entry;
            if (_taskPanes.ContainsKey(window))
            {
                entry = _taskPanes[window];
            }
            else
            {
                entry = new CTPManager<TTaskPane>.TaskPaneEntry();
                TTaskPane taskPane = _taskPaneFactory.CreateNewTaskPane(
                    window.Document, window);
                entry.UserTaskPane = taskPane;
                entry.MLTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(
                    taskPane, _taskPaneFactory.CreateTitle(taskPane), window);
                entry.MLTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                //Width of Pane
                entry.MLTaskPane.Width = 400;
                entry.MLTaskPane.VisibleChanged += new EventHandler(VstoTaskPane_VisibleChanged);
                _taskPanes.Add(window, entry);
            }
            entry.MLTaskPane.Visible = !entry.MLTaskPane.Visible;
            RefreshToggleButtons();
            return entry.UserTaskPane;
        }

        internal IEnumerable<TTaskPane> GetTaskPanes()
        {
            List<TTaskPane> taskPanes = new List<TTaskPane>();
            foreach (TaskPaneEntry entry in _taskPanes.Values)
            {
                taskPanes.Add(entry.UserTaskPane);
            }
            return taskPanes;
        }

        void Application_WindowActivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            RefreshToggleButtons();
        }

        void RefreshToggleButtons()
        {
            foreach (RibbonToggleButton button in _managedButtons)
            {
                button.Checked = IsTaskPaneVisible(Globals.ThisAddIn.Application.ActiveWindow);
            }
        }

        void VstoTaskPane_VisibleChanged(object sender, EventArgs e)
        {
            Tools.CustomTaskPane taskPane = (Tools.CustomTaskPane)sender;
            TTaskPane userTaskPane = _taskPanes[(MSWord.Window)taskPane.Window].UserTaskPane;
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
            TTaskPane CreateNewTaskPane(MSWord.Document document, MSWord.Window window);
        }

        struct TaskPaneEntry
        {
            public TTaskPane UserTaskPane;
            public Tools.CustomTaskPane MLTaskPane;
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