using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Nl.Infotron.Parsing;


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
            Microsoft.Office.Interop.Excel.Application X = Application;

            foreach (Excel.Worksheet w in Application.Worksheets)
            {
                foreach (Excel.Range r in w.Cells)
                {
                    string Formula = r.Formula;

                    //couple istest from core here
                    ExcelFormulaParser P = new ExcelFormulaParser();
                    if (P.IsTestFormula(Formula))
                    {
                        r.Interior.Color = 24;
                    }

                }
            }

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
