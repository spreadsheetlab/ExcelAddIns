using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace Polaris
{
    class AnalyzedCell
    {
        public struct PrecedentCell: IEquatable <PrecedentCell>
        {
            public string ID;
            public Excel.Range Cell;
            public Excel.Range Dependent;
            public int Level;

            public PrecedentCell(Excel.Range c, int l, Excel.Range dependent)
            {
                ID = c.Parent.Name + "!" + c.Address;
                Cell = c;
                Level = l;
                Dependent = dependent;
            }
            public bool Equals(PrecedentCell other)
            {
                return this.ID == other.ID;
            }
        }
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
                PFormulaAnalyzer fa = new PFormulaAnalyzer(c.Cell.Formula);
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
                    uniquePrecedents.Add(pc.ID, pc.Cell.Address);
                    cellsToProcess.Enqueue(pc);
                }
                transitivePrecedents.Add(pc);
                //System.Diagnostics.Debug.WriteLine("Add direct precedent " + pc.cell.Address + " for cell " + thisCell.Address + " with formula " + thisCell.Formula, "Polaris");
                System.Diagnostics.Debug.WriteLine("Add direct precedent " + pc.ID + " for cell " + thisCell.Address + " with formula " + (string)thisCell.Formula,"Polaris");
                PrintPrecedentsQueue(cellsToProcess);
            }
            while (cellsToProcess.Count > 0)
            {
                PrecedentCell pc = cellsToProcess.Dequeue();
                System.Diagnostics.Debug.WriteLine("Dequeue " + pc.Cell.Address, "Polaris");
                PrintPrecedentsQueue(cellsToProcess);
                int level = ++pc.Level;
                PCell cellToProcess = new PCell(pc.Cell);
                foreach (Excel.Range p in cellToProcess.Precedents)
                {
                    PrecedentCell precedent = new PrecedentCell(p, level, pc.Cell);
                    if (!uniquePrecedents.ContainsKey(precedent.ID))
                    {
                        uniquePrecedents.Add(precedent.ID, precedent.Cell.Address);
                        cellsToProcess.Enqueue(new PrecedentCell(p, level, pc.Cell));
                    }
                    System.Diagnostics.Debug.WriteLine("Add new precedent, level " + level.ToString() + ", " + p.Address + " for cell " + pc.Cell.Address + " with formula " + (string)pc.Cell.Formula, "Polaris");
                    PrintPrecedentsQueue(cellsToProcess);
                    transitivePrecedents.Add(new PrecedentCell(p, level, pc.Cell));
                }
            }
            return transitivePrecedents;
        }
        private void PrintPrecedentsQueue(Queue<PrecedentCell> queue)
        {
            foreach (PrecedentCell pc in queue)
            {
                Debug.Indent();
                Debug.WriteLine(pc.Cell.Address + '|' + pc.Level.ToString(), "Content queue");
                Debug.Unindent();
            }
        }
    }
}
