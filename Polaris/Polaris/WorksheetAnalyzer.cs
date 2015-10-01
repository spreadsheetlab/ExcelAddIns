using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using NLog;

namespace Polaris
{
    class WorksheetAnalyzer : IDisposable
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        private Excel.Worksheet thisSheet;
        private Excel.Workbook wkb;
        private Dictionary<string, Excel.Range> uniqueFormulas;
        public List<OutputCell> OutputCellsWithFunctions
        {
            get
            {
                return GetOutputCellsWithFunctions();
            }
        }
        public WorksheetAnalyzer(Excel.Worksheet wks)
        {
            thisSheet = wks;
            wkb = wks.Parent;
        }
        private Dictionary<string,Excel.Range> GetUniqueFormulas(Excel.Range rng)
        {
            uniqueFormulas = new Dictionary<string, Excel.Range>();
            var areas = rng.Areas;
            for (int i = 1; i <= areas.Count; i++)
            {
                Excel.Range r = areas[i];
                var rows = r.Rows;
                var columns = r.Columns;
                for (int row = 1; row <= rows.Count; row++)
                {
                    for (int c = 1; c <= columns.Count; c++)
                    {
                        if (!uniqueFormulas.ContainsKey(r[row,c].FormulaR1C1))
                        {
                            uniqueFormulas.Add(r[row, c].FormulaR1C1, r[row, c]);
                        }
                    }
                }
                Marshal.FinalReleaseComObject(r);
            }
            Marshal.FinalReleaseComObject(areas);
            return uniqueFormulas;
        }
        private bool isOutput(Excel.Range rng)
        {
            // Display dependent arrows for rng
            rng.ShowDependents();
            // Navigate arrow
            Excel.Range dependent = rng.NavigateArrow(false, 1, 1);
            // If there are no dependants the dependent address (target of arrow) is equal to the address of the selected cell (source of arrow).
            // No dependants means that the selecte cell is an output cell.
            bool isOutputCell;
            if (dependent.Address == rng.Address)
            {
                isOutputCell = true;
            }
            else
            {
                isOutputCell = false;
            }
            Marshal.FinalReleaseComObject(dependent);
            return isOutputCell;
        }
        private List<Excel.Range> GetOutputCells()
        {
            List<Excel.Range> outputCells = new List<Excel.Range>();

            var usedRange = thisSheet.UsedRange;
            var allFormulas = usedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
            uniqueFormulas = GetUniqueFormulas(allFormulas);

            foreach (KeyValuePair<string, Excel.Range> f in uniqueFormulas)
            {
                if (isOutput(f.Value))
                {
                    outputCells.Add(f.Value);
                }
            }
            Marshal.FinalReleaseComObject(usedRange);
            Marshal.FinalReleaseComObject(allFormulas);
            return outputCells;
        }
        private List<Excel.Range> GetDirectPrecedents(Excel.Range cell)
        {
            int arrowNr = 0;
            bool isNewArrow;
            Excel.Range precedent;
            Dictionary<string, Excel.Range> uniqueFormulas = new Dictionary<string, Excel.Range>();
            List<Excel.Range>precedents = new List<Excel.Range>();

            cell.ShowPrecedents();

            do
            {
                ++arrowNr;
                isNewArrow = true;
                int linkNr = 0;
                do
                {
                    ++linkNr;
                    try
                    {
                        precedent = cell.NavigateArrow(true, arrowNr, linkNr);
                    }
                    catch (Exception e)
                    {
                        break;
                    }
                    if (precedent.Address == cell.Address)
                    {
                        break;
                    }
                    else
                    {
                        isNewArrow = false;
                        if (precedent.Count == 1)
                        {
                            if (precedent.HasFormula)
                            {
                                precedents.Add(precedent);
                            }
                        }
                        else
                        {
                            try
                            {
                                var precedentFormulas = precedent.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                                uniqueFormulas = GetUniqueFormulas(precedentFormulas);
                                foreach (Excel.Range uniqueFormula in uniqueFormulas.Values)
                                {
                                    precedents.Add(uniqueFormula);
                                }
                            }
                            catch (Exception )
                            {
                            }
                        }
                    }
                } while (true);
                if (isNewArrow)
                {
                    break;
                }
            } while (true);
            return precedents;
        }
        private List<Excel.Range> GetAllPrecedents(Excel.Range cell)
        {
            List<Excel.Range> transitivePrecedents = new List<Excel.Range>();
            Queue<Excel.Range> cellsToProcess = new Queue<Excel.Range>();
            HashSet<string> uniquePrecedents = new HashSet<string>();
            List<Excel.Range> directPrecedents = GetDirectPrecedents(cell);
            foreach (Excel.Range p in directPrecedents)
            {
                if (uniquePrecedents.Add(thisSheet.Name + "!" + p.Address))
                {
                    cellsToProcess.Enqueue(p);
                }
                transitivePrecedents.Add(p);
            }
            while (cellsToProcess.Count > 0)
            {
                Excel.Range pc = cellsToProcess.Dequeue();
                directPrecedents = GetDirectPrecedents(pc);
                foreach (Excel.Range p in directPrecedents)
                {
                    if (uniquePrecedents.Add(thisSheet.Name + "!" + p.Address))
                    {
                        cellsToProcess.Enqueue(p);
                    }
                    transitivePrecedents.Add(p);
                }
            }
            return transitivePrecedents;
        }
        private List<string> GetAllFunctions(Excel.Range cell)
        {
            List<string>functions = new List<string>();
            HashSet<string> uniqueFunctions = new HashSet<string>();
            List<Excel.Range> transitivePrecedents = GetAllPrecedents(cell);

            foreach (Excel.Range c in transitivePrecedents)
            {
                try
                {
                    PFormulaAnalyzer fa = new PFormulaAnalyzer(c.Formula);
                    foreach (string function in fa.BuiltinFunctions())
                    {
                        uniqueFunctions.Add(function);
                    }
                }
                catch (Exception e)
                {
                    logger.Error(e.Message);
                }
            }
            return uniqueFunctions.ToList<string>();
        }
        private List<OutputCell> GetOutputCellsWithFunctions()
        {
            List<OutputCell> getOutputCellsWithFunctions = new List<OutputCell>();
            List<Excel.Range> outputCells = GetOutputCells();
            for (int i = 0; i < outputCells.Count ; i++)
            {
                List<string> functions = GetAllFunctions(outputCells[i]);
                if (functions.Count > 0)
                {
                    OutputCell cell = new OutputCell();
                    cell.CellAddress = outputCells[i].Address;
                    cell.Functions = functions;
                    cell.WorkbookName = wkb.Name;
                    cell.WorksheetName = thisSheet.Name;
                    getOutputCellsWithFunctions.Add(cell);
                }
            }
            return getOutputCellsWithFunctions;
        }
        public void Dispose()
        {
            if (uniqueFormulas != null)
            {
                foreach (var f in uniqueFormulas)
                {
                    Marshal.FinalReleaseComObject(f.Value);
                }
                uniqueFormulas = null;
            }
        }
    }
}
