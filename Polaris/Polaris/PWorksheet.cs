using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Polaris
{
    class PWorksheet
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
            wks.ClearArrows();
        }

        private Dictionary<string, Excel.Range> setUniqueFormulas()
        {
            // Get all formulas
            Excel.Range allFormulas = wks.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
            uniqueFormulas = new Dictionary<string, Excel.Range>();
            foreach (Excel.Range c in allFormulas)
            {
                // Only add formulas to dictionary if the R1C1 formula is unique
                if (!uniqueFormulas.ContainsKey(c.FormulaR1C1))
                {
                    uniqueFormulas.Add(c.FormulaR1C1, c);
                }
            }
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
    }
}
