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
                return uniqueFormulas;
            }
        }
        public Dictionary<string, Excel.Range> OutputCells
        {
            get
            {
                return outputCells;
            }
        }


        public PWorksheet(Excel.Worksheet worksheet)
        {
            wks = worksheet;
            setUniqueFormulas();
            setOutputCells(uniqueFormulas);
            wks.ClearArrows();
        }

        private void setUniqueFormulas()
        {
            Excel.Range allFormulas = wks.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
            uniqueFormulas = new Dictionary<string, Excel.Range>();
            foreach (Excel.Range c in allFormulas)
            {
                if (!uniqueFormulas.ContainsKey(c.FormulaR1C1))
                {
                    uniqueFormulas.Add(c.FormulaR1C1, c);
                }
            }
        }

        private void setOutputCells(Dictionary<string, Excel.Range> uniqueFormulas)
        {
            outputCells = new Dictionary<string, Excel.Range>();

            foreach (KeyValuePair<string, Excel.Range> f in uniqueFormulas)
            {
                if (isOutput(f.Value))
                {
                    outputCells.Add(f.Key, f.Value);
                }
            }
        }
        private bool isOutput(Excel.Range rng)
        {
            rng.ShowDependents();
            Excel.Range dependent = rng.NavigateArrow(false, 1, 1);
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
