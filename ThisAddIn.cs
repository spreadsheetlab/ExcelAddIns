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
        List<string> TestFormulas = new List<string>();
        VerifyTests V;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        public void FindTests()
        {
            Excel.Worksheet w = Application.ActiveWorkbook.ActiveSheet;

            V = new VerifyTests(this);

            foreach (Excel.Range cell in w.UsedRange)
            {
                if (cell.HasFormula)
                {
                    string Formula = cell.Formula.Substring(1, cell.Formula.Length - 1);
                        
                    ExcelFormulaParser P = new ExcelFormulaParser();

                    if (P.IsTestFormula(Formula))
                    {
                        int Passing = P.GetPassingBranch(Formula);

                        string Text = P.GetCondition(Formula) + " should be ";

                        if (P.GetPassingBranch(Formula) == 0) //if the first branch is passing, the test codintion should be true, else false
                        {
                            Text += "true";
                        }
                        else
                        {
                            Text += "false";
                        }

                        V.PrintTest(Text, Formula);                                           
                    }
                }
            }

            V.Show();

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

        internal void MarkTests()
        {
            if (TestFormulas.Count == 0)
            {
                MessageBox.Show("No tests saved yet, first run 'save tests'");
            }
            else
            {
                foreach (string item in TestFormulas)
                {
                    
                }
            }
        }

        internal void SaveTests(List <string> formulas)
        {
            TestFormulas = formulas;
        }
    }
}
