using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Polaris
{
    class PWorksheet : IDisposable
    {
        private Excel.Worksheet wks;
        private Dictionary<string, Excel.Range> uniqueFormulas;
        private Dictionary<string, Excel.Range> outputCells;

        public Dictionary<string,Excel.Range> UniqueFormulas
        {
            get
            {
                return uniqueFormulas ?? (uniqueFormulas = this.setUniqueFormulas());
            }
        }
        public Dictionary<string, Excel.Range> OutputCells
        {
            get
            {
                return outputCells ?? (outputCells = this.setOutputCells(UniqueFormulas));
            }
        }


        public PWorksheet(Excel.Worksheet worksheet)
        {
            wks = worksheet;
        }

        private Dictionary<string, Excel.Range> setUniqueFormulas()
        {
            try
            {
                var xlUsedRange = wks.UsedRange;
                var allFormulas = xlUsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                uniqueFormulas = new Dictionary<string, Excel.Range>();
                for (int a = 1; a <= allFormulas.Areas.Count; a++)
                {
                    Excel.Range rng = allFormulas.Areas[a];
                    for (int r = 1; r <= rng.Rows.Count; r++)
                    {
                        for (int c = 1; c <= rng.Columns.Count; c++)
                        {
                            if(!uniqueFormulas.ContainsKey(rng[r,c].FormulaR1C1))
                            {
                                uniqueFormulas.Add(rng[r, c].FormulaR1C1, rng[r, c]);
                            }
                        }
                    }
                }
                Marshal.FinalReleaseComObject(xlUsedRange);
                Marshal.FinalReleaseComObject(allFormulas);
            }
            catch (Exception)
            {
                uniqueFormulas = new Dictionary<string, Excel.Range>();
            }
            // Get all formulas
            return uniqueFormulas;
        }

        private Dictionary<string, Excel.Range> setOutputCells(Dictionary<string, Excel.Range> uniqueFormulas)
        {
            
            outputCells = new Dictionary<string, Excel.Range>();

            foreach (KeyValuePair<string, Excel.Range> f in uniqueFormulas)
            {
                if (isOutput(f.Value))
                {
                    outputCells.Add(f.Key, f.Value);
                }
            }
            return outputCells;
        }
        private bool isOutput(Excel.Range rng)
        {
            // Display dependent arrows for rng
            rng.ShowDependents();
            // Navigate arrow
            Excel.Range dependent = rng.NavigateArrow(false, 1, 1);
            // If there are no dependants the dependent address (target of arrow) is equal to the address of the selected cell (source of arrow).
            // No dependants means that the selecte cell is an output cell.
            if (dependent.Address == rng.Address)
            {
                return true;
            }
            else
            {
                return false;
            }
            
        }
        public void Dispose()
        {
            foreach (var rng in outputCells)
            {
                Marshal.FinalReleaseComObject(rng.Value);
            }
            foreach (var f in uniqueFormulas)
            {
                Marshal.FinalReleaseComObject(f.Value);
            }
            outputCells = null;
            uniqueFormulas = null;
        }
    }
}
