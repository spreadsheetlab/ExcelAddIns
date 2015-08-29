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

namespace Polaris
{
    class PolarisController
    {
        public void CreateGraphOutput(AnalyzedCell cell)
        {
            List<SEdge<string>> edges = new List<SEdge<string>>();
            foreach (AnalyzedCell.PrecedentCell p in cell.Precedents)
            {
                edges.Add(new SEdge<string>(p.ID, p.Dependent.Parent.Name + "!" + p.Dependent.Address));
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
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            string folderToScan = getFolder();
            string[] files = Directory.GetFiles(folderToScan, "*.*", SearchOption.TopDirectoryOnly);
            int row = 0;
            foreach (string f in files)
            {
                Excel.Workbook outputWorkbook = Globals.ThisAddIn.Application.Workbooks.Add();
                Excel.Worksheet outputSheet = outputWorkbook.Worksheets[1];
                Excel.Workbook analyzedWorkbook = Globals.ThisAddIn.Application.Workbooks.Open(Filename:f, ReadOnly: true);
                foreach (Excel.Worksheet wks in analyzedWorkbook.Worksheets)
                {
                    Globals.ThisAddIn.Application.StatusBar = wks.Name;
                    PWorksheet currentSheet = new PWorksheet(wks);
                    Dictionary<string, Excel.Range> outputCells = currentSheet.OutputCells;
                    foreach (KeyValuePair<string,Excel.Range> c in outputCells)
                    {
                        ++row;
                        outputSheet.get_Range("A" + Convert.ToString(row)).Value = wks.Name;
                        outputSheet.get_Range("B" + Convert.ToString(row)).Value = c.Value.Address;
                        outputSheet.get_Range("C" + Convert.ToString(row)).Value = "'" + c.Value.Formula;
                        AnalyzedCell oCell = new AnalyzedCell(c.Value);
                        int precedentColumn = 0;
                        foreach (AnalyzedCell.PrecedentCell p in oCell.Precedents)
                        { 
                            outputSheet.get_Range("D" + Convert.ToString(row)).Offset[0,precedentColumn].Value = "'" + p.Level + "|" + p.Cell.Address;
                            ++precedentColumn;
                        }
                    }
                }
            }
            Globals.ThisAddIn.Application.StatusBar = false;
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        public void AnalyseSingleCell(Excel.Range cell)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Excel.Workbook outputWorkbook = Globals.ThisAddIn.Application.Workbooks.Add();
            Excel.Worksheet outputSheet = outputWorkbook.Worksheets[1];
            AnalyzedCell oCell = new AnalyzedCell(cell);
            int row = 1;
            foreach (AnalyzedCell.PrecedentCell p in oCell.Precedents)
            {
                outputSheet.get_Range("A" + Convert.ToString(row)).Value = "'" + p.Level + "|" + p.Dependent.Address + "|" + p.ID;
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
