namespace GettingStartedOutlookAddIn
{
    partial class DemoRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public DemoRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonDemo = this.Factory.CreateRibbonButton();
            this.button2Demo = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
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
            this.group1.Items.Add(this.buttonDemo);
            this.group1.Items.Add(this.button2Demo);
            this.group1.Label = "Demo";
            this.group1.Name = "group1";
            // 
            // buttonDemo
            // 
            this.buttonDemo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonDemo.Label = "Get Title";
            this.buttonDemo.Name = "buttonDemo";
            this.buttonDemo.ShowImage = true;
            this.buttonDemo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDemo_Click);
            // 
            // button2Demo
            // 
            this.button2Demo.Label = "Get text selected";
            this.button2Demo.Name = "button2Demo";
            this.button2Demo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2Demo_Click);
            // 
            // DemoRibbon
            // 
            this.Name = "DemoRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.DemoRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDemo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2Demo;
    }

    partial class ThisRibbonCollection
    {
        internal DemoRibbon DemoRibbon
        {
            get { return this.GetRibbon<DemoRibbon>(); }
        }
    }
}
