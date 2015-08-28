namespace Polaris
{
    partial class PolarisRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public PolarisRibbon()
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
            this.polarisRibbonTab = this.Factory.CreateRibbonTab();
            this.polarisTabGroup1 = this.Factory.CreateRibbonGroup();
            this.buttonStart = this.Factory.CreateRibbonButton();
            this.buttonAnalyseCell = this.Factory.CreateRibbonButton();
            this.buttonGraph = this.Factory.CreateRibbonButton();
            this.polarisRibbonTab.SuspendLayout();
            this.polarisTabGroup1.SuspendLayout();
            // 
            // polarisRibbonTab
            // 
            this.polarisRibbonTab.Groups.Add(this.polarisTabGroup1);
            this.polarisRibbonTab.Label = "Polaris";
            this.polarisRibbonTab.Name = "polarisRibbonTab";
            // 
            // polarisTabGroup1
            // 
            this.polarisTabGroup1.Items.Add(this.buttonStart);
            this.polarisTabGroup1.Items.Add(this.buttonAnalyseCell);
            this.polarisTabGroup1.Items.Add(this.buttonGraph);
            this.polarisTabGroup1.Label = "Analysis";
            this.polarisTabGroup1.Name = "polarisTabGroup1";
            // 
            // buttonStart
            // 
            this.buttonStart.Label = "Start Analysis";
            this.buttonStart.Name = "buttonStart";
            // 
            // buttonAnalyseCell
            // 
            this.buttonAnalyseCell.Label = "Analyse Selected Cell";
            this.buttonAnalyseCell.Name = "buttonAnalyseCell";
            this.buttonAnalyseCell.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAnnalyseCell_Click);
            // 
            // buttonGraph
            // 
            this.buttonGraph.Label = "Test Graph";
            this.buttonGraph.Name = "buttonGraph";
            this.buttonGraph.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGraph_Click_1);
            // 
            // PolarisRibbon
            // 
            this.Name = "PolarisRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.polarisRibbonTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Polaris_Load);
            this.polarisRibbonTab.ResumeLayout(false);
            this.polarisRibbonTab.PerformLayout();
            this.polarisTabGroup1.ResumeLayout(false);
            this.polarisTabGroup1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab polarisRibbonTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup polarisTabGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAnalyseCell;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGraph;
    }

    partial class ThisRibbonCollection
    {
        internal PolarisRibbon Polaris
        {
            get { return this.GetRibbon<PolarisRibbon>(); }
        }
    }
}
