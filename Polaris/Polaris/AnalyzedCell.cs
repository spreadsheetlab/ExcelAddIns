using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Polaris
{
    class AnalyzedCell : IDisposable
    {
        private int worksheetID;
        private string spreadsheetID;
        private string formula;
        private string formulaR1C1;
        private List<PrecedentCell> transitivePrecedents;
        private Excel.Range thisCell;
        private Dictionary<string,string> uniquePrecedents;
        private List<string> functions;

        public int WorksheetID { get { return worksheetID; } }
        public string SpreadsheetID { get { return spreadsheetID; } }
        public string Formula { get { return formula; } }
        public string FormulaR1C1 { get { return formulaR1C1; } }
        public List<PrecedentCell> TransitivePrecedents 
        { 
            get 
            {
                return transitivePrecedents ?? (transitivePrecedents = this.GetAllPrecedents());
            } 
        }
        public List<string> Functions
        {
            get
            {
                return functions ?? (functions = this.GetAllFunctions());
            }
        }
        public AnalyzedCell(Excel.Range c)
        {
            thisCell = c;
            worksheetID = c.Parent.Index;
            spreadsheetID = c.Parent.Parent.Name;
            formulaR1C1 = c.FormulaR1C1;
            formula = c.Formula;
        }
        private List<string> GetAllFunctions()
        {
            functions = new List<string>();
            HashSet<string> uniqueFunctions = new HashSet<string>();
            if (transitivePrecedents == null) GetAllPrecedents();
            
            foreach (PrecedentCell c in transitivePrecedents)
            {
                PFormulaAnalyzer fa = new PFormulaAnalyzer(c.ThisCell.Formula);
                foreach (string formula in fa.BuiltinFunctions())
                {
                    uniqueFunctions.Add(formula);
                }
            }
            return uniqueFunctions.ToList<string>();
        }
        private List<PrecedentCell> GetAllPrecedents()
        {
            transitivePrecedents = new List<PrecedentCell>();
            Queue<PrecedentCell> cellsToProcess = new Queue<PrecedentCell>();
            uniquePrecedents = new Dictionary<string, string>();
            PCell c = new PCell(thisCell);
            foreach (Excel.Range p in c.Precedents)
            {
                PrecedentCell pc = new PrecedentCell(p, 1, thisCell);
                if (!uniquePrecedents.ContainsKey(pc.ID))
                {
                    uniquePrecedents.Add(pc.ID, pc.ThisCell.Address);
                    cellsToProcess.Enqueue(pc);
                }
                transitivePrecedents.Add(pc);
            }
            while (cellsToProcess.Count > 0)
            {
                PrecedentCell pc = cellsToProcess.Dequeue();
                int level = ++pc.Level;
                PCell cellToProcess = new PCell(pc.ThisCell);
                foreach (Excel.Range p in cellToProcess.Precedents)
                {
                    PrecedentCell precedent = new PrecedentCell(p, level, pc.ThisCell);
                    if (!uniquePrecedents.ContainsKey(precedent.ID))
                    {
                        uniquePrecedents.Add(precedent.ID, precedent.ThisCell.Address);
                        cellsToProcess.Enqueue(new PrecedentCell(p, level, pc.ThisCell));
                    }
                    transitivePrecedents.Add(new PrecedentCell(p, level, pc.ThisCell));
                }
            }
            return transitivePrecedents;
        }
        public void Dispose()
        {
            foreach (PrecedentCell pc  in transitivePrecedents)
            {
                Marshal.FinalReleaseComObject(pc.ThisCell);
                Marshal.FinalReleaseComObject(pc.DependentCell);
            }
            transitivePrecedents = null;
        }
    }
}
