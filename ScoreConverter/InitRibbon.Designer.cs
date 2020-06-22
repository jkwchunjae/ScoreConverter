namespace ScoreConverter
{
    partial class InitRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public InitRibbon()
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
            this.InitRibbonTab = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.OpenFormButton = this.Factory.CreateRibbonButton();
            this.InitRibbonTab.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // InitRibbonTab
            // 
            this.InitRibbonTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.InitRibbonTab.Groups.Add(this.group2);
            this.InitRibbonTab.Label = "TabAddIns";
            this.InitRibbonTab.Name = "InitRibbonTab";
            // 
            // group2
            // 
            this.group2.Items.Add(this.OpenFormButton);
            this.group2.Label = "group2";
            this.group2.Name = "group2";
            // 
            // OpenFormButton
            // 
            this.OpenFormButton.Label = "Start";
            this.OpenFormButton.Name = "OpenFormButton";
            this.OpenFormButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenFormButton_Click);
            // 
            // InitRibbon
            // 
            this.Name = "InitRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.InitRibbonTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.InitRibbon_Load);
            this.InitRibbonTab.ResumeLayout(false);
            this.InitRibbonTab.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab InitRibbonTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton OpenFormButton;
    }

    partial class ThisRibbonCollection
    {
        internal InitRibbon InitRibbon
        {
            get { return this.GetRibbon<InitRibbon>(); }
        }
    }
}
