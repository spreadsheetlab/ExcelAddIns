namespace Expector
{
    partial class ExpectorRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExpectorRibbon()
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
            this.ExpectorTab = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.FindTestButton = this.Factory.CreateRibbonButton();
            this.MarkTestButton = this.Factory.CreateRibbonButton();
            this.ColorTestsButton = this.Factory.CreateRibbonButton();
            this.MarkTestedButton = this.Factory.CreateRibbonButton();
            this.MakeNonTestButton = this.Factory.CreateRibbonButton();
            this.coverageButton = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.addTestSmelly = this.Factory.CreateRibbonButton();
            this.addTestBig = this.Factory.CreateRibbonButton();
            this.addTestReferences = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.ExpectorTab.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
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
            // ExpectorTab
            // 
            this.ExpectorTab.Groups.Add(this.group2);
            this.ExpectorTab.Groups.Add(this.group3);
            this.ExpectorTab.Groups.Add(this.group4);
            this.ExpectorTab.Label = "Expector";
            this.ExpectorTab.Name = "ExpectorTab";
            // 
            // group2
            // 
            this.group2.Items.Add(this.FindTestButton);
            this.group2.Items.Add(this.MarkTestButton);
            this.group2.Items.Add(this.ColorTestsButton);
            this.group2.Label = "Find and run tests";
            this.group2.Name = "group2";
            // 
            // FindTestButton
            // 
            this.FindTestButton.Label = "Find existing tests in this spreadsheet";
            this.FindTestButton.Name = "FindTestButton";
            this.FindTestButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FindTestButton_Click);
            // 
            // MarkTestButton
            // 
            this.MarkTestButton.Label = "Run all tests";
            this.MarkTestButton.Name = "MarkTestButton";
            this.MarkTestButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RunTestButton_Click);
            // 
            // ColorTestsButton
            // 
            this.ColorTestsButton.Label = "Color tests (green = passing, red = failing)";
            this.ColorTestsButton.Name = "ColorTestsButton";
            this.ColorTestsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ColorTestsButtonClick);
            // 
            // MarkTestedButton
            // 
            this.MarkTestedButton.Label = "Show me what is tested";
            this.MarkTestedButton.Name = "MarkTestedButton";
            this.MarkTestedButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MarkTestedButton_Click);
            // 
            // MakeNonTestButton
            // 
            this.MakeNonTestButton.Label = "Show me what is not tested";
            this.MakeNonTestButton.Name = "MakeNonTestButton";
            this.MakeNonTestButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MakeNonTestButton_Click);
            // 
            // coverageButton
            // 
            this.coverageButton.Label = "How well is this spreadsheet tested?";
            this.coverageButton.Name = "coverageButton";
            this.coverageButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.coverageButton_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.MarkTestedButton);
            this.group3.Items.Add(this.MakeNonTestButton);
            this.group3.Items.Add(this.coverageButton);
            this.group3.Label = "Understand quality of tests";
            this.group3.Name = "group3";
            // 
            // group4
            // 
            this.group4.Items.Add(this.addTestSmelly);
            this.group4.Items.Add(this.addTestBig);
            this.group4.Items.Add(this.addTestReferences);
            this.group4.Label = "Add new tests";
            this.group4.Name = "group4";
            // 
            // addTestSmelly
            // 
            this.addTestSmelly.Label = "I want to test a complex formula";
            this.addTestSmelly.Name = "addTestSmelly";
            this.addTestSmelly.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addTestSmelly_Click);
            // 
            // addTestBig
            // 
            this.addTestBig.Label = "I want to test a formula with a big value";
            this.addTestBig.Name = "addTestBig";
            // 
            // addTestReferences
            // 
            this.addTestReferences.Label = "I want to test a formula with many references";
            this.addTestReferences.Name = "addTestReferences";
            // 
            // ExpectorRibbon
            // 
            this.Name = "ExpectorRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.ExpectorTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.ExpectorTab.ResumeLayout(false);
            this.ExpectorTab.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab ExpectorTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FindTestButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MarkTestButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MarkTestedButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MakeNonTestButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ColorTestsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton coverageButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addTestSmelly;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addTestBig;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addTestReferences;
    }

    partial class ThisRibbonCollection
    {
        internal ExpectorRibbon Ribbon1
        {
            get { return this.GetRibbon<ExpectorRibbon>(); }
        }
    }
}
