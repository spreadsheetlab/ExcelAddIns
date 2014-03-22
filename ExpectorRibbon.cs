using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace Expector
{
    public partial class ThisAddIn //default name is ExpectorRibbon and then you CANNOT connect to Application variable so a rename is needed
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void FindTestButton_Click(object sender, RibbonControlEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application X = Application; 

            foreach (Excel.Worksheet w in Application.Worksheets)
            {
                foreach (Excel.Range r in w.Cells)
                {
                    string Formula = r.Formula;

                    //couple istest from core here
                }
            }

        }
    }
}
