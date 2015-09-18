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
        private int maxId = 1;
        private int functionId = 1;
        private Dictionary<string, int> functions;
        public OutputGenerator()
        {
            DeleteOutputFiles();
        }
        public void AddOutputCell(Excel.Range cell)
        {
            var engine = new FileHelperEngine<CSV_OutputCell>();
            CSV_OutputCell outputCell = new CSV_OutputCell();
            outputCell.Id = maxId;
            outputCell.WorkbookName = cell.Parent.Parent.Name;
            outputCell.WorksheetName = cell.Parent.Name;
            outputCell.CellAddress = cell.Address;
            engine.AppendToFile("OutputCells.txt", outputCell);
            ++maxId;
        }
        public void AddTransaction(AnalyzedCell cell)
        {
            var engine = new FileHelperEngine<CSV_Transaction>();
            CSV_Transaction transaction = new CSV_Transaction();
            List<string> functionNumbers = new List<string>();
            foreach (string f in cell.Functions)
            {
                functionNumbers.Add(functions[f].ToString());
            }
            transaction.functions = string.Join(" ", functionNumbers);
            engine.AppendToFile("Transactions.txt", transaction);

        }
        public void AddFunctions(AnalyzedCell cell)
        {
            var engine = new FileHelperEngine<CSV_Function>();
            if (functions == null) functions = new Dictionary<string, int>();
            foreach (string f in cell.Functions)
            {
                if (! functions.ContainsKey(f))
                {
                    functions.Add(f, functionId);
                    CSV_Function csv_function = new CSV_Function();
                    csv_function.Id = functionId;
                    csv_function.Function = f;
                    engine.AppendToFile("Functions.txt", csv_function);
                    ++functionId;
                }
            }
        }

        private void DeleteOutputFiles()
        {
            string[] files = {"OutputCells.txt","Functions.txt","Transactions.txt"};
            foreach (string f in files)
            {
                if (File.Exists(f)) File.Delete(f);
            }
        }

    }
}
