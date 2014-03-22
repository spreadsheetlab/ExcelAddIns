﻿namespace Expector
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
            this.MarkCoverageButton = this.Factory.CreateRibbonButton();
            this.MakeNonTestButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.ExpectorTab.SuspendLayout();
            this.group2.SuspendLayout();
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
            this.ExpectorTab.Label = "Expector";
            this.ExpectorTab.Name = "ExpectorTab";
            // 
            // group2
            // 
            this.group2.Items.Add(this.FindTestButton);
            this.group2.Items.Add(this.MarkTestButton);
            this.group2.Items.Add(this.MarkCoverageButton);
            this.group2.Items.Add(this.MakeNonTestButton);
            this.group2.Name = "group2";
            // 
            // FindTestButton
            // 
            this.FindTestButton.Label = "Find Tests";
            this.FindTestButton.Name = "FindTestButton";
            this.FindTestButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FindTestButton_Click);
            // 
            // MarkTestButton
            // 
            this.MarkTestButton.Label = "Mark Tests";
            this.MarkTestButton.Name = "MarkTestButton";
            // 
            // MarkCoverageButton
            // 
            this.MarkCoverageButton.Label = "Mark Covered Formulas";
            this.MarkCoverageButton.Name = "MarkCoverageButton";
            // 
            // MakeNonTestButton
            // 
            this.MakeNonTestButton.Label = "Mark Non-Covered Formulas";
            this.MakeNonTestButton.Name = "MakeNonTestButton";
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

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab ExpectorTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FindTestButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MarkTestButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MarkCoverageButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton MakeNonTestButton;
    }

    partial class ThisRibbonCollection
    {
        internal ExpectorRibbon Ribbon1
        {
            get { return this.GetRibbon<ExpectorRibbon>(); }
        }
    }
}
