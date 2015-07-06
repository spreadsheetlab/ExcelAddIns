using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace Expector
{
    public partial class ExpectorRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void RunTestButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Expector.RunTests();
        }

        private void FindTestButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Expector.InitializeTests();
        }

        private void ColorTestsButtonClick(object sender, RibbonControlEventArgs e)
        {
            Globals.Expector.ColorTests();
        }

        private void MarkTestedButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Expector.HighLightTested();
        }

        private void MakeNonTestButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Expector.HighLightNonTested();
        }

        private void coverageButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Expector.Coverage();
        }

        private void addTestSmelly_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.Expector.ProposeSmellyCell();
        }
    }
}
