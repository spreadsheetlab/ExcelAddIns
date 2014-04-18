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

        private void MarkTestButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Test();
        }

        private void FindTestButton_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
