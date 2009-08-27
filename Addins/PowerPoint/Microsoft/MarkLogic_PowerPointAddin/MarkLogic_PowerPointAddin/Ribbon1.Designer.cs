namespace MarkLogic_PowerPointAddin
{
    partial class Ribbon1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.group1 = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.tab2 = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.group2 = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.viewTaskPaneButton = new Microsoft.Office.Tools.Ribbon.RibbonToggleButton();
            this.button2 = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.button3 = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group2);
            this.tab2.Label = "Mark Logic";
            this.tab2.Name = "tab2";
            // 
            // group2
            // 
            this.group2.Items.Add(this.viewTaskPaneButton);
            this.group2.Label = "authoring";
            this.group2.Name = "group2";
            // 
            // viewTaskPaneButton
            // 
            this.viewTaskPaneButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.viewTaskPaneButton.Label = "enable kit";
            this.viewTaskPaneButton.Name = "viewTaskPaneButton";
            this.viewTaskPaneButton.ShowImage = true;
            this.viewTaskPaneButton.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.toggleButton1_Click);
            // 
            // button2
            // 
            this.button2.Image = global::MarkLogic_PowerPointAddin.Properties.Resources.menu1_Image;
            this.button2.Label = "Save To MarkLogic";
            this.button2.Name = "button2";
            this.button2.Position = Microsoft.Office.Tools.Ribbon.RibbonPosition.AfterOfficeId("FileSaveAsMenu");
            this.button2.ShowImage = true;
            this.button2.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Image = global::MarkLogic_PowerPointAddin.Properties.Resources.menu1_Image;
            this.button3.Label = "Save To MarkLogic As";
            this.button3.Name = "button3";
            this.button3.Position = Microsoft.Office.Tools.Ribbon.RibbonPosition.AfterOfficeId("FileSaveAsMenu");
            this.button3.ShowImage = true;
            this.button3.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.button3_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            // 
            // Ribbon1.OfficeMenu
            // 
            this.OfficeMenu.Items.Add(this.button3);
            this.OfficeMenu.Items.Add(this.button2);
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton viewTaskPaneButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
