using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace AragornAddIn
{
    public partial class Ribbon1
    {


        


        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TurnOnAragorn();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {

            Globals.ThisAddIn.TurnOffAragorn();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ProcessWorkBook();

            //button3.Enabled = false;
            button1.Enabled = true;
        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }


    }
}
