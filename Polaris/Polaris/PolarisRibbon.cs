using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace Polaris
{
    public partial class PolarisRibbon
    {
        private void Polaris_Load(object sender, RibbonUIEventArgs e)
        {
            this.buttonStart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(buttonStart_Click);
        }

        private void buttonStart_Click(object sender, RibbonControlEventArgs e)
        {
            PolarisController pc = new PolarisController();
            pc.startAnalysis();
        }

        private void buttonAnnalyseCell_Click(object sender, RibbonControlEventArgs e)
        {
            PolarisController pc = new PolarisController();
            pc.AnalyseSingleCell(Globals.ThisAddIn.Application.ActiveCell);
        }
    }
}
