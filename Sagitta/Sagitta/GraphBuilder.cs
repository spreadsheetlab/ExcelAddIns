using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NetOffice;
using NetOffice.Tools;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.ExcelApi.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.OfficeApi.Tools;
using VBIDE = NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Enums;
using System.Windows.Forms;
using XLParser;
using Irony.Parsing;

namespace Sagitta
{
    class GraphBuilder
    {
        public void BuildGraphFromWorkbook(Excel.Workbook wkb)
        {
            // Loop trough all the worksheets
            Excel.Sheets allWks = wkb.Worksheets;
            foreach (Excel.Worksheet wks in allWks)
            {
                // Select all formulas on the worksheet
                Excel.Range allFormulas = wks.Cells.SpecialCells(XlCellType.xlCellTypeFormulas);
                // Loop trough all formulas
                int counter = 0;
                foreach (Excel.Range formula in allFormulas)
                {
                    ++counter;
                    // Get References from formula
                    IEnumerable<ParseTreeNode> refNodes = ExcelFormulaParser.GetReferenceNodes(ExcelFormulaParser.Parse(formula.Formula.ToString()));
                    foreach (ParseTreeNode node in refNodes)
                    {
                        MessageBox.Show(node.ChildNodes[0].ChildNodes[0].Token.Text);
                    }
                }
                MessageBox.Show(counter.ToString());
            }
        }
    }
}
