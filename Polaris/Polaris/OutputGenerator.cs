using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using FileHelpers;
using System.IO;
using NLog;

namespace Polaris
{
    class OutputGenerator
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        private int maxId = 1;
        private int functionId = 1;
        private Dictionary<string, int> functions;
        public OutputGenerator()
        {
            //DeleteOutputFiles();
        }
        public void AddOutputCell(Excel.Range cell)
        {
            var engine = new FileHelperEngine<CSV_OutputCell>();
            CSV_OutputCell outputCell = new CSV_OutputCell();
            outputCell.WorkbookName = cell.Parent.Parent.Name;
            outputCell.WorksheetName = cell.Parent.Name;
            outputCell.CellAddress = cell.Address;
            engine.AppendToFile("OutputCells.txt", outputCell);
            ++maxId;
        }
        public void WriteOutputAndTransactionToFile(List<OutputCell> cells)
        {
            
            var engineOutputCells = new FileHelperEngine<CSV_OutputCell>();
            var engineTransactions = new FileHelperEngine<CSV_Transaction>();
            List<CSV_OutputCell> outputCells = new List<CSV_OutputCell>();
            List<CSV_Transaction> transactions = new List<CSV_Transaction>();
            List<string> convertedFunctions = new List<string>();
            foreach (var c in cells)
            {
                try
                {
                    CSV_OutputCell outputCell = new CSV_OutputCell();
                    CSV_Transaction transaction = new CSV_Transaction();
                    outputCell.CellAddress = c.CellAddress;
                    outputCell.WorkbookName = c.WorkbookName;
                    outputCell.WorksheetName = c.WorksheetName;
                    outputCells.Add(outputCell);
                    convertedFunctions = functionIntegers(c.Functions);
                    transaction.functions = string.Join(" ", convertedFunctions);
                    transactions.Add(transaction);
                }
                catch (Exception e)
                {
                    logger.Error(e.Message);
                }

            }
            engineOutputCells.AppendToFile("OutputCells.txt", outputCells);
            engineTransactions.AppendToFile("Transactions.txt", transactions);
        }
        private List<string> functionIntegers(List<string> functions)
        {
            List<string> excelFunctions = Polaris.Properties.Resources.ExcelFunctions.Split(new string[] { "\n" }, StringSplitOptions.None).ToList<string>();
            List<string> functionIntegers = new List<string>();
            foreach (string function in functions)
            {
                int index = excelFunctions.IndexOf(function);
                if (index != -1)
                {
                    functionIntegers.Add(index.ToString());
                }
                else
                {
                    throw  new InvalidOperationException("Function " + function + " does not exist in functionlist");
                }
            }
            return functionIntegers;
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
