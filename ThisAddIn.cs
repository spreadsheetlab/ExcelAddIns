using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Infotron.Parsing;
using System.Windows.Forms;


namespace Expector
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void Test()
        {
            Excel.Worksheet w = Application.ActiveWorkbook.ActiveSheet;

            int nTests = 0;

                foreach (Excel.Range cell in w.UsedRange)
                {
                    if (cell.HasFormula)
                    {
                        string Formula = cell.Formula.Substring(1, cell.Formula.Length - 1);
                        

                        ExcelFormulaParser P = new ExcelFormulaParser();

                        if (P.IsTestFormula(Formula))
                        {
                            cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                            nTests++;
                        }
                    }
                }

                MessageBox.Show(String.Format("{0} tests detected", nTests.ToString()));
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
