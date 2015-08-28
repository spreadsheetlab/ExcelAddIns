using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace Polaris
{
    class OutputCell
    {
        public struct ProcessedCell: IEquatable <ProcessedCell>
        {
            public string ID;
            public Excel.Range Cell;
            public Excel.Range Dependent;
            public int Level;

            public ProcessedCell(Excel.Range c, int l, Excel.Range dependent)
            {
                ID = c.Parent.Name + "!" + c.Address;
                Cell = c;
                Level = l;
                Dependent = dependent;
            }
            public bool Equals(ProcessedCell other)
            {
                return this.ID == other.ID;
            }
        }
        private int worksheetID;
        private string spreadsheetID;
        private string formula;
        private string formulaR1C1;
        private List<ProcessedCell> precedents;
        private Excel.Range thisCell;
        private Dictionary<string,string> uniquePrecedents;

        public int WorksheetID { get { return worksheetID; } }
        public string SpreadsheetID { get { return spreadsheetID; } }
        public string Formula { get { return formula; } }
        public string FormulaR1C1 { get { return formulaR1C1; } }
        public List<ProcessedCell> Precedents { get { return precedents; } }
        public OutputCell(Excel.Range c)
        {
            thisCell = c;
            worksheetID = c.Parent.Index;
            spreadsheetID = c.Parent.Parent.Name;
            formulaR1C1 = c.FormulaR1C1;
            formula = c.Formula;
            precedents = new List<ProcessedCell>();
            GetAllPrecedents();
        }
        private void GetAllPrecedents()
        {
            Queue<ProcessedCell> cellsToProcess = new Queue<ProcessedCell>();
            uniquePrecedents = new Dictionary<string, string>();
            PCell c = new PCell(thisCell);
            foreach (Excel.Range p in c.Precedents)
            {
                ProcessedCell pc = new ProcessedCell(p, 1, thisCell);
                if (!uniquePrecedents.ContainsKey(pc.ID))
                {
                    uniquePrecedents.Add(pc.ID, pc.Cell.Address);
                    cellsToProcess.Enqueue(pc);
                }
                precedents.Add(pc);
                //System.Diagnostics.Debug.WriteLine("Add direct precedent " + pc.cell.Address + " for cell " + thisCell.Address + " with formula " + thisCell.Formula, "Polaris");
                System.Diagnostics.Debug.WriteLine("Add direct precedent " + pc.ID + " for cell " + thisCell.Address + " with formula " + (string)thisCell.Formula,"Polaris");
                PrintPrecedentsQueue(cellsToProcess);
            }
            while (cellsToProcess.Count > 0)
            {
                ProcessedCell pc = cellsToProcess.Dequeue();
                System.Diagnostics.Debug.WriteLine("Dequeue " + pc.Cell.Address, "Polaris");
                PrintPrecedentsQueue(cellsToProcess);
                int level = ++pc.Level;
                PCell cellToProcess = new PCell(pc.Cell);
                foreach (Excel.Range p in cellToProcess.Precedents)
                {
                    ProcessedCell precedent = new ProcessedCell(p, level, pc.Cell);
                    if (!uniquePrecedents.ContainsKey(precedent.ID))
                    {
                        uniquePrecedents.Add(precedent.ID, precedent.Cell.Address);
                        cellsToProcess.Enqueue(new ProcessedCell(p, level, pc.Cell));
                    }
                    System.Diagnostics.Debug.WriteLine("Add new precedent, level " + level.ToString() + ", " + p.Address + " for cell " + pc.Cell.Address + " with formula " + (string)pc.Cell.Formula, "Polaris");
                    PrintPrecedentsQueue(cellsToProcess);
                    precedents.Add(new ProcessedCell(p, level, pc.Cell));
                }
            }
        }
        private void PrintPrecedentsQueue(Queue<ProcessedCell> queue)
        {
            foreach (ProcessedCell pc in queue)
            {
                Debug.Indent();
                Debug.WriteLine(pc.Cell.Address + '|' + pc.Level.ToString(), "Content queue");
                Debug.Unindent();
            }
        }
    }
}
