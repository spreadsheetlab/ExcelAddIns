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
using Infotron.Util;


namespace Expector
{
    public class testFormula
    {
        public string original;
        public string condition;
        public bool shouldbe;
        public string worksheet;
        public string location;
    }

    public partial class Expector
    {
        List<testFormula> TestFormulas = new List<testFormula>();
        VerifyTests V;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        public void InitializeTests()
        {
            V = new VerifyTests(this);
            int ntests = 0;

            foreach (Excel.Worksheet w in Application.ActiveWorkbook.Worksheets)
            {
                //we limit the number of cells to analyze to 250, otherwise it will be too slow.
                int AnalyzedCells = 0;
                foreach (Excel.Range cell in w.UsedRange)
                {
                    AnalyzedCells++;
                    if (AnalyzedCells > 250)
                    {
                        continue;
                    }
                    if (cell.HasFormula)
                    {
                        try
                        {
                            testFormula t = new testFormula();
                            t.original = cell.Formula.Substring(1, cell.Formula.Length - 1);
                            t.location = cell.AddressLocal.Replace("$", "");
                            t.worksheet = w.Name;

                            ExcelFormulaParser P = new ExcelFormulaParser();

                            if (P.IsTestFormula(t.original))
                            {
                                ntests++;
                                int Passing = P.GetPassingBranch(t.original);

                                string Condition = P.GetCondition(t.original);

                                List<String> ConditionList = P.Split(Condition);

                                string Text;

                                if (ConditionList.Count == 1) //just 1 item in the condition, like IF(A4,"OK","NOT OK")
                                {
                                    Text = Condition + " should ";

                                    if (P.GetPassingBranch(t.original) == 0) //if the first branch is passing, the test condition should be true, else false
                                    //the senmatic of IF(A4) is that it is false if A4 = 0, true in all other cases.
                                    {
                                        Text += "not be 0"; //it should be true, so non-zero
                                        t.shouldbe = true;
                                    }
                                    else
                                    {
                                        Text += "be 0";
                                        t.shouldbe = false;
                                    }
                                }
                                else
                                {
                                    if (ConditionList.Count == 2) //a function as condition 1, like IF(ISBLANK(A4),"OK","NOT OK")
                                    {
                                        Text = Condition + " should ";
                                        if (P.GetPassingBranch(t.original) == 0) //if the first branch is passing, the test condition should be true, else false                                    
                                        {
                                            Text += "hold"; //it should be true, so non-zero
                                            t.shouldbe = true;
                                        }
                                        else
                                        {
                                            Text += "not hold";
                                            t.shouldbe = false;
                                        }
                                    }

                                    else //the condition is a 'real' condition i.e. IF(A5 > 5, "OK", "ERROR")
                                    {
                                        Text = ConditionList[0];

                                        if (P.GetPassingBranch(t.original) == 0) //if the first branch is passing, the test codintion should be true, else false
                                        {
                                            Text += " should be " + ConditionList[1] + ConditionList[2]; //this adds the > 5 part
                                            t.shouldbe = true;
                                        }
                                        else
                                        {
                                            Text += " should be " + Invert(ConditionList[1]) + ConditionList[2]; //this adds the <= 5 part
                                            t.shouldbe = false;
                                        }
                                    }
                                }

                                V.PrintTest(Text, t);
                            }
                        }
                        catch (Exception)
                        {
                            //just skip this cell
                        }
                       
                    }
                }
            }
            if (ntests == 0)
            {
                MessageBox.Show("No tests found");
            }
            else
            {
                V.Show();
            }
            
        }

        private string Invert(string p)
        {
            if (p == ">") return "<=";
            if (p == "<") return ">=";
            if (p == ">=") return "<";
            if (p == "<=") return ">";
            if (p == "<>") return "=";
            if (p == "=") return "<>";

            return "not p";
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


        private Excel.Worksheet GetWorksheetByName(string name)
        {
            foreach (Excel.Worksheet worksheet in Application.ActiveWorkbook.Worksheets)
            {
                if (worksheet.Name == name)
                {
                    return worksheet;
                }
            }
            throw new ArgumentException();
        }

        internal void ColorTests()
        {
            if (TestFormulas.Count == 0)
            {
                MessageBox.Show("No tests saved yet, first run 'Initialize tests'");
            }
            else
            {
                Excel.Worksheet w = GetWorksheetByName("Expector-Tests");
                int ntests = w.UsedRange.Rows.Count;

                for (int i = 1; i <= ntests; i++)
                {
                    //get the tests value:
                    var result = w.Cells.Item[i, 1].value;

                    bool bool_result = GetBool(result);

                    Excel.Range testCell = GetTestatRowi(w, i);

                    if (bool_result)
                    {
                        testCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                    }
                    else
                    {
                        testCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                }
               
            }
        }

        private Excel.Range GetTestatRowi(Excel.Worksheet w, int i)
        {
            //get the location of the test
            Excel.Worksheet testSheet = GetWorksheetByName(w.Cells.Item[i, 2].value);
            Location L = new Location(w.Cells.Item[i, 3].value);

            //get the cell
            Excel.Range testCell = testSheet.Cells.Item[L.Row + 1, L.Column + 1];
            return testCell;
        }

        private static bool GetBool(dynamic result)
        {
            //a condition can either be a boolean or a double
            //if it is a boolean, return it
            //if it is a double, we return whether it is equal to 0. (0 == false, all other values are true)
            bool bool_result;
            if (!(result is bool))
            {
                bool_result = (result != 0d);
            }
            else
            {
                bool_result = result;
            }
            return bool_result;
        }

        internal void SaveTests(List <testFormula> formulas)
        {
            TestFormulas = formulas;
            int i = 1;

            foreach (var item in TestFormulas)
            {
                //is there already a worksheet to save tests in?
                Excel.Worksheet w;

                try
                {
                    w = GetWorksheetByName("Expector-Tests");
                }
                catch (SystemException E)
                {
                    //w is not found

                    //get the last worksheet to add Expector-Tests at the end
                    Excel.Worksheet Last = this.Application.Worksheets.get_Item(this.Application.Worksheets.Count);
                    w = (Excel.Worksheet)this.Application.Worksheets.Add(missing,Last);
                    w.Name = "Expector-Tests";                             
                }

                ExcelFormulaParser P = new ExcelFormulaParser();

                
                string formula = P.GetCondition(item.original, item.worksheet);

                if (item.shouldbe == false)
                {
                    formula = "NOT(" + formula + ")";
                }

                w.Cells.Item[i, 1].Formula = "=" + formula;

                //print worksheet
                w.Cells.Item[i, 2].Value = item.worksheet;
                w.Cells.Item[i, 3].Value = item.location;

                i++;
                
            }

        }

        internal void RunTests()
        {
            if (TestFormulas.Count == 0)
            {
                MessageBox.Show("No tests saved yet, first run 'Initialize tests'");
            }
            else
            {
                Excel.Worksheet w = GetWorksheetByName("Expector-Tests");
                int ntests = w.UsedRange.Rows.Count;
                string toPrint = "";

                for (int i = 1; i <= ntests; i++)
                {
                    //get the tests value:
                    var result = w.Cells.Item[i, 1].value;

                    bool bool_result=GetBool(result);

                    //get the tests condition:
                    string formula = w.Cells.Item[i, 1].formula;

                    List<String> testResults = new List<string>();

                    if (bool_result)
                    {
                        testResults.Add("Test passed " + formula); 
                    }
                    else
                    {
                        testResults.Add("Test failed " + formula); 
                    }


                    foreach (var item in testResults)
                    {
                        toPrint = toPrint += item + Environment.NewLine;
                    }

                }
                MessageBox.Show(toPrint);

            }
        }

        internal void HighLightTested()
        {
            if (TestFormulas.Count == 0)
            {
                MessageBox.Show("No tests saved yet, first run 'Initialize tests'");
            }
            else
            {
                ResetCellColors();

                List<Excel.Range> cellsToColor = GetCoveredCells();

                foreach (Excel.Range prec in cellsToColor)
                {
                    prec.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                }
            }
        }

        private List<Excel.Range> GetCoveredCells()
        {

            Excel.Worksheet w = GetWorksheetByName("Expector-Tests");

            List<Excel.Range> cellsToColor = new List<Excel.Range>();
            int ntests = w.UsedRange.Rows.Count;
            for (int i = 1; i <= ntests; i++)
            {
                //for eacht test, get all precedents and color them

                Excel.Range testCell = GetTestatRowi(w, i);
                Excel.Range precs;

                try
                {
                    precs = testCell.Precedents; //unfortunately, there is no hasprecedents propoerty...
                }
                catch (Exception e)
                {
                    // no precedents found
                    precs = null;
                }

                if (precs != null)
                {
                    foreach (Excel.Range item in precs)
                    {
                        if (!ContainsCell(cellsToColor, item))
                        {
                            cellsToColor.Add(item);
                        }
                    }

                    if (!ContainsCell(cellsToColor, testCell))
                    {
                        cellsToColor.Add(testCell);
                    }
                }

            }
            return cellsToColor;
        }

        private static bool ContainsCell(List<Excel.Range> list, Excel.Range item)
        {
            //we cannot use the normal .contains, it does not work on COM objects because they are copies.
            bool found = false;

            foreach (Excel.Range item2 in list)
            {
                if ((item2.Address == item.Address) && (item2.Worksheet.Name == item.Worksheet.Name))
                {
                    found = true;
                }
            }
            return found;
        }


        internal void HighLightNonTested()
        {
            List<Excel.Range> coveredCells = GetCoveredCells();

            ResetCellColors();

            foreach (Excel.Range Cell in Application.ActiveWorkbook.ActiveSheet.UsedRange)
            {
                if (!ContainsCell(coveredCells, Cell))
                    {
                        Cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    }
            }

        }

        private void ResetCellColors()
        {
            foreach (Excel.Range Cell in Application.ActiveWorkbook.ActiveSheet.UsedRange)
            {
                Cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Transparent);
            }
        }
    }
}
