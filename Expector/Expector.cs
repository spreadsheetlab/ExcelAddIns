using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using GemBox.Spreadsheet;
using Infotron.Parsing;
using Infotron.PerfectXL.DataModel;
using Infotron.PerfectXL.SmellAnalyzer;
using Infotron.PerfectXL.SmellAnalyzer.SmellAnalyzer;
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
        public AnalysisController controller;
        public List<testFormula> TestFormulas = new List<testFormula>();


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
            this.Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpenHandle);
        }

        #endregion

        private void Application_WorkbookOpenHandle(Excel.Workbook Wb)
        {
            //try to find the sheet where the tests are loaded. If found load, if not do nothing

            try
            {
                Excel.Worksheet w = GetWorksheetByName("Expector-Tests");
    
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
                    
                    TestFormulas.Add(f);
                }
            }
            catch (Exception)
            {
                //no problem, maybe the user will init the tests this time.
            }



        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {


            
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


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

        internal void SaveTests()
        {
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
            if (TestFormulas.Count == 0)
            {
                MessageBox.Show("No tests saved yet, first run 'Initialize tests'");
            }
            else
            {
                Excel.Worksheet w = GetWorksheetByName("Expector-Tests");
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

        internal void HighLightTested()
        {
            if (TestFormulas.Count == 0)
            {
                MessageBox.Show("No tests saved yet, first run 'Initialize tests'");
            }
            else
            {
                ResetCellColors();

                List<Excel.Range> cellsToColor = GetCoveredCells(false);

                foreach (Excel.Range prec in cellsToColor)
                {
                    prec.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                }
            }
        }

        private List<Excel.Range> GetCoveredCells(Boolean Global)
        {
            //if global is false, only cells on the current worksheet (activesheet) are returned

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
                        
                        if (Global || item.Worksheet == (Excel.Worksheet)this.Application.ActiveSheet)
                        {
                            if (!ContainsCell(cellsToColor, item))
                            {
                                cellsToColor.Add(item);
                            }
                        }

                    }
                }


                if (Global || testCell.Worksheet == (Excel.Worksheet)this.Application.ActiveSheet)
                {
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
            List<Excel.Range> coveredCells = GetCoveredCells(false);

            ResetCellColors();

            foreach (Excel.Range Cell in Application.ActiveWorkbook.ActiveSheet.UsedRange)
            {
                if (!ContainsCell(coveredCells, Cell) && Cell.Value != null)
                    {
                        Cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    }
            }
        }

        private void ResetCellColors()
        {
            foreach (Excel.Range Cell in Application.ActiveWorkbook.ActiveSheet.UsedRange)
            {
                Cell.Interior.ColorIndex = 0;
            }
        }

        internal void Coverage()
        {
            List<Excel.Range> coveredCells = GetCoveredCells(true);

            List<Excel.Range> allFormulas = new List<Excel.Range>() ;
            //get all cells:

            foreach (Excel.Worksheet w in Application.ActiveWorkbook.Worksheets)
            {
                foreach (Excel.Range c in w.UsedRange)
                {
                    if (c.Value2 != null)
	                {
                        allFormulas.Add(c);
	                }
                }
            }

            double coverage = (double)coveredCells.Count / allFormulas.Count;

            coverage = coverage * 100;

            string message = String.Format("{0}% of all non-empty cells are covered by at least one test", Math.Round(coverage));

            MessageBox.Show(message);
        }

        internal void ProposeSmellyCell()
        {
            SpreadsheetInfo.SetLicense("E7OS-D3IG-PM8L-A03O");

            if (!Application.ActiveWorkbook.Saved)
            {
                Application.Dialogs[Excel.XlBuiltInDialog.xlDialogSaveAs].Show();
            }

            if (!Application.ActiveWorkbook.Saved)
            {
                MessageBox.Show("The workbook must be saved before analysis. Aborting.");
                return;
            }

            controller = new AnalysisController
            {
                Worker = new BackgroundWorker { WorkerReportsProgress = true },
                Filename = Application.ActiveWorkbook.FullName,
            };


            controller.RiskAnalyzers = new List<IRiskAnalyzer>()
            {
                new ManyOperationsAnalyzer(),
                new ManyReferencesAnalyzer(),
            };

            controller.RunAnalysis();

            if (!controller.Spreadsheet.AnalysisSucceeded)
            {
                throw new Exception(controller.Spreadsheet.ErrorMessage);
            }


            //detected smells now contains all smells
            List<Excel.Range> coveredCells = GetCoveredCells(true);

            List<string> testedCellLocations = new List<string>();
            foreach (Excel.Range c in coveredCells)
            {
                string loc = c.Worksheet.Name+"!"+c.Address.Replace("$","");
                testedCellLocations.Add(loc);
            }

            //first filter out only the high risk smells:
            List<Smell> sortedSmells = controller.DetectedSmells.Where(x => x.RiskValue > 4).OrderBy(x => x.RiskValue).ToList();
            
            //then, find the corresponding cells:
            List<Cell> smellyCells = sortedSmells.Select(x => ((SiblingClass)x.Source).Cells.First()).ToList();

            //now locate the non-covered ones
            smellyCells = smellyCells.Where(x => !testedCellLocations.Contains(x.GetLocationIncludingSheetnameString())).ToList();

            if (smellyCells.Count() > 0)
            {
                Cell maxCell = smellyCells.First();
                string message = String.Format("You could a a test for the cell on {0}: {1}. Do you want to do this?", maxCell.GetLocationIncludingSheetnameString(), maxCell.Formula);

                DialogResult result1 = MessageBox.Show(message, "Add new test", MessageBoxButtons.YesNo);

                if (result1 == DialogResult.Yes)
                {
                    var A = new AddTest(this, maxCell);
                    A.Show();
                }
            }
            else
            {
                //there are no complex cells to test, for now do nothing
                MessageBox.Show("No complex formulas found to test, hooray!");
            }
        }

        internal void ProposeHighCoverageCell()
        {

        }

        internal void ProposeLargeCell()
        {
            List<Excel.Range> nonCoveredCells = getNonCoveredCells();
            
            float maxvalue = int.MinValue;
            Excel.Range maxCell = nonCoveredCells[0];

            foreach (Excel.Range cell in nonCoveredCells)
            {
                try
                {
                    float v = (float)cell.Value;
                    if (v > maxvalue)
                    {
                        maxvalue = v;
                        maxCell = cell;
                    }
                }
                catch (Exception)
                {

                }


            }

            MessageBox.Show(String.Format("What about {0} with value {1}?", maxCell.AddressLocal, maxvalue));
        }

        private List<Excel.Range> getNonCoveredCells()
        {
            List<Excel.Range> coveredCells = GetCoveredCells(true);

            List<Excel.Range> nonCoveredCells = new List<Excel.Range>();

            foreach (Excel.Worksheet w in Application.ActiveWorkbook.Worksheets)
            {
                if (w.Name != "Expector-Tests")
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
    }


}
