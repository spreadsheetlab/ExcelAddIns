using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using XLParser;
using Irony.Parsing;
using QuickGraph;
using QuickGraph.Serialization;
using QuickGraph.Algorithms;
using QuickGraph.Collections;
using QuickGraph.Contracts;
using QuickGraph.Data;
using QuickGraph.Graphviz;
using QuickGraph.Predicates;
using System.Xml;
using NLog;
using System.Runtime.InteropServices;

namespace Polaris
{
    class PolarisController
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        public void CreateGraphOutput(AnalyzedCell cell)
        {
            List<SEdge<string>> edges = new List<SEdge<string>>();
            foreach (PrecedentCell p in cell.TransitivePrecedents)
            {
                edges.Add(new SEdge<string>(p.ID, p.DependentCell.Parent.Name + "!" + p.DependentCell.Address));
            }
            var graph = edges.ToAdjacencyGraph<string, SEdge<string>>();
            using (var xw = XmlWriter.Create("test.graphml"))
            {
                graph.SerializeToGraphML(xw,graph.GetVertexIdentity(),graph.GetEdgeIdentity());
            }
            //string output = graph.ToGraphviz();
            var graphiz = new GraphvizAlgorithm<string, SEdge<string>>(graph);
            graphiz.FormatVertex += OnFormatVertex;
            var output = graphiz.Generate(new FileDotEngine(),"graphiz.txt");
            //File.WriteAllText("graphiz.txt", output);
        }
        public virtual void OnFormatVertex(object obj, FormatVertexEventArgs<string> v)
        {
            v.VertexFormatter.Label = v.Vertex.ToString();
        }
        public void startAnalysis()
        {
            Excel._Application xlApp = Globals.ThisAddIn.Application;
            xlApp.ScreenUpdating = false;
            string folderToScan = getFolder();
            string[] files = Directory.GetFiles(folderToScan, "*.*", SearchOption.TopDirectoryOnly);
            OutputGenerator outputGenerator = new OutputGenerator();
            int fileNumber = 0;
            
            foreach (string f in files)
            {
                ++fileNumber;
                xlApp.DisplayAlerts = false;
                Excel.Workbooks xlWorkbooks = xlApp.Workbooks;
                Excel.Workbook analyzedWorkbook = xlWorkbooks.Open(Filename:f, ReadOnly: true, UpdateLinks: 2);
                xlApp.DisplayAlerts = true;
                int wksNumber = 0;
                foreach (Excel.Worksheet wks in analyzedWorkbook.Worksheets)
                {
                    try
                    {
                        ++wksNumber;
                        using(PWorksheet currentSheet = new PWorksheet(wks))
                        {
                            Dictionary<string, Excel.Range> outputCells = currentSheet.OutputCells;
                            int cellNumber = 0;
                            foreach (KeyValuePair<string, Excel.Range> c in outputCells)
                            {
                                try
                                {
                                    ++cellNumber;
                                    Globals.ThisAddIn.Application.StatusBar = "File " + fileNumber + " of " + files.Count() + ", Sheet " + wksNumber + " of " + analyzedWorkbook.Worksheets.Count + ", Cell " + cellNumber + " of " + outputCells.Count();
                                    using (AnalyzedCell oCell = new AnalyzedCell(c.Value))
                                    {
                                        if (oCell.Functions.Count > 0)
                                        {
                                            outputGenerator.AddOutputCell(c.Value);
                                            outputGenerator.AddFunctions(oCell);
                                            outputGenerator.AddTransaction(oCell);
                                        }
                                    }
                                }
                                catch (Exception e)
                                {
                                    logger.Error("skipped cell|" + c.Value.Address + "|" +
                                        "|" + f + "|" + wks.Name +
                                        "|" + e.Message);
                                }

                           }
                        }
                    }
                    catch (Exception e)
                    {
                        logger.Error(e.Message);
                        //logger.Error("skipped worksheet|" + f + "|" + wks.Name + "|" + e.Message);
                    }
                    Marshal.FinalReleaseComObject(wks);
                }
                xlApp.DisplayAlerts = false;
                analyzedWorkbook.Close(false);
                xlApp.DisplayAlerts = true;
                Marshal.FinalReleaseComObject(xlWorkbooks);
                Marshal.FinalReleaseComObject(analyzedWorkbook);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            xlApp.StatusBar = false;
            xlApp.ScreenUpdating = true;
            //Marshal.FinalReleaseComObject(xlApp);
            //GC.Collect();
            //GC.WaitForPendingFinalizers();
        }

        public void AnalyseSingleCell(Excel.Range cell)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Excel.Workbook outputWorkbook = Globals.ThisAddIn.Application.Workbooks.Add();
            Excel.Worksheet outputSheet = outputWorkbook.Worksheets[1];
            AnalyzedCell oCell = new AnalyzedCell(cell);
            string test = oCell.Functions.ToString();
            int row = 1;
            foreach (PrecedentCell p in oCell.TransitivePrecedents)
            {
                outputSheet.get_Range("A" + Convert.ToString(row)).Value = "'" + p.Level + "|" + p.DependentCell.Address + "|" + p.ID;
                ++row;
            }
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            Debug.Flush();
            CreateGraphOutput(oCell);
        }
        public string getFolder()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.ShowDialog();
            return fbd.SelectedPath;
        }
    }
}
