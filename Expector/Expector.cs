using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Infotron.Parsing;
using Infotron.PerfectXL.DataModel;
using Infotron.PerfectXL.SmellAnalyzer;
using Infotron.Util;
using Excel = Microsoft.Office.Interop.Excel;


namespace Expector
{
    public class testFormula
    {
        public string original; 
        public string condition;
        public string worksheet;
        public string location;
    }

    public partial class Expector
    {
        public const string expectorWorksheetName = "Expector-Tests";
        public AnalysisController controller;
        public List<testFormula> testFormulas = new List<testFormula>();
        public List<Excel.Range> coveredCells = new List<Excel.Range>();
        public List<Excel.Range> nonCoveredCells = new List<Excel.Range>();


        #region initialization of Expector (load testformulas etc.)

        private void InternalStartup()
        {
            this.Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpenHandle);
        }

        public List<Excel.Range> GetCoveredCells()
        {
            Excel.Worksheet w = GetExpectorSheet();

            List<Excel.Range> coveredCells = new List<Excel.Range>();
            int ntests = w.UsedRange.Rows.Count;
            for (int i = 1; i <= ntests; i++)
            {
                //for eacht test, get all precedents and add them to the list of covered cells

                Excel.Range testCell = GetTestatRowi(w, i);

                if (testCell != null)
                {
                    //add the test cell
                    if (!ContainsCell(coveredCells, testCell))
                    {
                        coveredCells.Add(testCell);
                    }
                    

                    Excel.Range precs;

                    try
                    {
                        precs = testCell.Precedents; //unfortunately, there is no hasprecedents propoerty
                        //we do not need recursion though, because this is recursive already
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
                            if (item.Value != null) //empty cells are not considered covered
                            {
                                if (!ContainsCell(coveredCells, item))
                                {
                                    coveredCells.Add(item);
                                }   
                            }

                        }
                    }
                }
            }
            return coveredCells;
        }

        private List<Excel.Range> getNonCoveredCells()
        {
            List<Excel.Range> nonCoveredCells = new List<Excel.Range>();

            foreach (Excel.Worksheet w in Application.ActiveWorkbook.Worksheets)
            {
                if (w.Name != expectorWorksheetName)
                {
                    foreach (Excel.Range cell in w.UsedRange.Cells)
                    {
                        if (!ContainsCell(coveredCells, cell) && cell.Value != null && cell.HasFormula)
                        {
                            nonCoveredCells.Add(cell);
                        }
                    }
                }

            }
            return nonCoveredCells;
        }

        private void Application_WorkbookOpenHandle(Excel.Workbook Wb)
        {
            //try to find the sheet where the tests are loaded. If found load, if not do nothing

            try
            {
                Excel.Worksheet w = GetExpectorSheet();
    
                int ntests = w.UsedRange.Rows.Count;

                for (int i = 1; i <= ntests; i++)
                {
                    string formula = w.Cells.Item[i, 1].formula;
                    formula = formula.Substring(1);

                    testFormula f = new testFormula()
                    {
                        //we read the tests value from the worksheet

                        condition = formula,
                        worksheet = w.Cells.Item[i, 2].value,
                        location = w.Cells.Item[i, 3].value,
                    };
                    
                    testFormulas.Add(f);
                }

                coveredCells = GetCoveredCells();
            }
            catch (Exception)
            {
                //no problem, maybe the user will init the tests this time.
            }



        }


        #endregion

        public void InitializeTests()
        {
            var V = new VerifyTests(this);
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
                            //is this formula a test formula?
                             ExcelFormulaParser P = new ExcelFormulaParser();

                            string formula = cell.Formula.Substring(1, cell.Formula.Length - 1);

                            if (P.IsTestFormula(formula))
                            {
                                testFormula t = new testFormula();
                                t.original = formula;
                                t.location = cell.AddressLocal.Replace("$", "");
                                t.worksheet = w.Name;

                                ntests++;

                                V.PrintTest(t);
                                
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

        private Excel.Worksheet GetExpectorSheet()
        {
            return GetWorksheetByName(expectorWorksheetName);
        }

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
            if (testFormulas.Count == 0)
            {
                MessageBox.Show("No tests saved yet, either extract or add them");
            }
            else
            {
                Excel.Worksheet w = GetExpectorSheet();
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
            Excel.Range worksheetCell = w.Cells.Item[i, 2];
            if (worksheetCell.Value != null)
            {
                //get the location of the test
                Excel.Worksheet testSheet = GetWorksheetByName(w.Cells.Item[i, 2].value);
                Location L = new Location(w.Cells.Item[i, 3].value);

                //get the cell
                Excel.Range testCell = testSheet.Cells.Item[L.Row + 1, L.Column + 1];
                return testCell;
            }
            else
            {
                return null;
            }

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

        internal void SaveTests()
        {
            int i = 1;

            foreach (var item in testFormulas)
            {
                //is there already a worksheet to save tests in?
                Excel.Worksheet w;

                try
                {
                    w = GetExpectorSheet();
                }
                catch (SystemException E)
                {
                    //w is not found

                    //get the last worksheet to add Expector-Tests at the end
                    Excel.Worksheet Last = this.Application.Worksheets.get_Item(this.Application.Worksheets.Count);
                    w = (Excel.Worksheet)this.Application.Worksheets.Add(missing,Last);
                    w.Name = expectorWorksheetName;                           
                }

                w.Cells.Item[i, 1].formula = "="+item.condition;
                w.Cells.Item[i, 2].Value = item.worksheet;
                w.Cells.Item[i, 3].Value = item.location;

                Excel.Range rangeToHoldHyperlink = w.get_Range(new Location(3, i-1).ToString(), Type.Missing);
                string hyperlinkTargetAddress = item.worksheet + "!" + item.location;
                w.Hyperlinks.Add(rangeToHoldHyperlink, string.Empty, hyperlinkTargetAddress, "", item.worksheet +"!" + item.location);

                i++;
                
            }

        }

        internal void RunTests()
        {
            if (testFormulas.Count == 0)
            {
                MessageBox.Show("No tests saved yet, either extract or add them");
            }
            else
            {
                Excel.Worksheet w = GetExpectorSheet();
                int ntests = w.UsedRange.Rows.Count;
                string toPrint = "";

                //we make two lists, one for the passed and one for the failed tests
                List<String> failingTests = new List<string>();
                List<String> passingTests = new List<string>();

                for (int i = 1; i <= ntests; i++)
                {
                    //get the tests value:
                    var result = w.Cells.Item[i, 1].value;

                    bool bool_result=GetBool(result);

                    //get the tests condition:
                    string formula = w.Cells.Item[i, 1].formula;

                    if (bool_result)
                    {
                        passingTests.Add(formula); 
                    }
                    else
                    {
                        failingTests.Add(formula); 
                    }

                }

                toPrint = "Tests passed: (" + passingTests.Count.ToString() + "/" + ntests.ToString() +")";
                toPrint += Environment.NewLine;

                foreach (var item in passingTests)
                {
                    toPrint = toPrint += item + Environment.NewLine;
                }

                toPrint += Environment.NewLine;

                toPrint += "Tests failed: (" + failingTests.Count.ToString() + "/" + ntests.ToString() + ")"; 
                toPrint += Environment.NewLine;

                foreach (var item in failingTests)
                {
                    toPrint = toPrint += item + Environment.NewLine;
                }

                MessageBox.Show(toPrint);

            }
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



        private void ResetCellColors()
        {
            foreach (Excel.Range Cell in Application.ActiveWorkbook.ActiveSheet.UsedRange)
            {
                Cell.Interior.ColorIndex = 0;
            }
        }

        internal void HighLightTested()
        {
            if (testFormulas.Count == 0)
            {
                MessageBox.Show("No tests saved yet, either extract or add them");
            }
            else
            {
                ResetCellColors();

                foreach (Excel.Range prec in coveredCells)
                {
                    prec.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                }
            }
        }



        internal void HighLightNonTested()
        {
            

            ResetCellColors();

            foreach (Excel.Range Cell in Application.ActiveWorkbook.ActiveSheet.UsedRange)
            {
                if (!ContainsCell(coveredCells, Cell) && Cell.Value != null) //could we use the nonCovered cells here?
                    {
                        Cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    }
            }
        }

        public double getCurrentCoverage()
        {
            List<Excel.Range> allNonEmptyCells = new List<Excel.Range>();

            foreach (Excel.Worksheet w in Application.ActiveWorkbook.Worksheets)
            {
                if (w.Name != expectorWorksheetName)
                {
                    foreach (Excel.Range c in w.UsedRange)
                    {
                        if (c.Value2 != null)
                        {
                            allNonEmptyCells.Add(c);
                        }
                    }
                }

            }

            double coverage = (double)coveredCells.Count / allNonEmptyCells.Count;

            coverage = coverage * 100;
            return coverage;
        }



        internal void ShowCoverage()
        {

            double coverage = getCurrentCoverage();

            string message = String.Format("{0}% of all non-empty cells are covered by at least one test", Math.Round(coverage));

            MessageBox.Show(message);
        }


        #region Propose Cells to add tests for



        public int getComplexity(Excel.Range c)
        {
            ExcelFormulaParser p = new ExcelFormulaParser();
            FormulaAnalyzer f = new FormulaAnalyzer(c.Formula.Substring(1), p);
            int complexity = f.References().Count + f.GetFunctions().Count;
            return complexity;
        }


        public delegate int CellMaxFunction(Excel.Range c);

        internal void ShowTestforMaxCellforGivenFunction(List<Excel.Range> nonCoveredCells, CellMaxFunction cellFunc)
        {
            int maxReferences = 0;
            Excel.Range maxCell = nonCoveredCells[0];

            foreach (Excel.Range c in nonCoveredCells)
            {
                int functionResult = cellFunc(c);

                if (functionResult > maxReferences)
                {
                    maxReferences = functionResult;
                    maxCell = c;
                }
            }

            //put focus on the smelly cell
            maxCell.Worksheet.Select();
            maxCell.Select();
            maxCell.Activate();  //TODO: nodig?              

            double coverageBefore = getCurrentCoverage();

            var A = new AddTest(this, maxCell.Worksheet.Name, maxCell.Formula, maxCell.Address.Replace("$", ""));
            A.Show();



        }



        internal void ProposeLargeCell()
        {
            List<Excel.Range> nonCoveredCells = getNonCoveredCells();

            if (nonCoveredCells.Count() > 0)
            {
                ShowTestforMaxCellforGivenFunction(nonCoveredCells, x => (int)x.Value);
            }
            else
            {
                //there are no complex cells to test, for now do nothing
                MessageBox.Show("No complex formulas found to test, hooray!");
            }
        }


        internal void ProposeSmellyCell()
        {
            List<Excel.Range> nonCoveredCells = getNonCoveredCells();

            if (nonCoveredCells.Count() > 0)
            {
                ShowTestforMaxCellforGivenFunction(nonCoveredCells, getComplexity);
            }
            else
            {
                //there are no complex cells to test, for now do nothing
                MessageBox.Show("No complex formulas found to test, hooray!");
            }
        }

        internal void ProposeHighCoverageCell()
        {
            List<Excel.Range> nonCoveredCells = getNonCoveredCells();

            if (nonCoveredCells.Count() > 0)
            {
                ShowTestforMaxCellforGivenFunction(nonCoveredCells, x => x.Precedents.Count);
            }
            else
            {
                //there are no complex cells to test, for now do nothing
                MessageBox.Show("No complex formulas found to test, hooray!");
            }
        }


        #endregion


    }


}
