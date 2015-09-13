using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using FileHelpers;
using System.IO;

namespace Polaris
{
    class OutputGenerator
    {
        private List<OutputCell> outputCells;
        private int maxId = 1;
        public OutputGenerator()
        {
            if (File.Exists("OutputCells.txt"))
            {
                File.Delete("OutputCells.txt");
            }
        }
        public void AddOutputCell(Excel.Range cell)
        {
            if (outputCells == null) outputCells = new List<OutputCell>();
            OutputCell outputCell = new OutputCell();
            outputCell.Id = maxId;
            outputCell.WorkbookName = cell.Parent.Parent.Name;
            outputCell.WorksheetName = cell.Parent.Name;
            outputCell.CellAddress = cell.Address;
            ++maxId;
            outputCells.Add(outputCell);
        }
        public void AppendToOutputCellFile()
        {
            var engine = new FileHelperEngine<OutputCell>();
            engine.AppendToFile("OutputCells.txt", outputCells);
        }

    }
}
