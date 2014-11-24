﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Collections;
using Microsoft.Office.Tools.Ribbon;
using Infotron.FSharpFormulaTransformation;
using Infotron.PerfectXL.SmellAnalyzer;
using System.ComponentModel;
using GemBox.Spreadsheet;
using PerfectXL.Domain.Observation;
using Infotron.PerfectXL.SmellAnalyzer.SmellAnalyzer;
using System.Drawing;
using Infotron.PerfectXL.DataModel;
using Infotron.Util;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelAddIn3
{
    public class TransformationComparer : System.Collections.IComparer
    {
        public int Compare(object x, object y)
        {

            if (((FSharpTransformationRule)x).priority == ((FSharpTransformationRule)y).priority)
                return 0;
            else if (((FSharpTransformationRule)x).priority > ((FSharpTransformationRule)y).priority)
                return 1;
            else
                return -1;
        }
    }

    public class SmellyCell
    {
        public Range Cell;
        public Object OriginalPattern;
        public Object OriginalColor;

        public SmellyCell(Range cell,
            Object originalPattern,
            Object originalColor)
        {
            this.Cell = cell;
            this.OriginalPattern = originalPattern;
            this.OriginalColor = originalColor;
        }

        public void Reset(){
            Cell.Interior.Color = OriginalColor;
            Cell.Interior.Pattern = OriginalPattern;
            Cell.Comment.Visible = false;
        }

        public void Apply(Smell smell)
        {
            Cell.Interior.Pattern = XlPattern.xlPatternSolid;
            Cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);

            if (Cell.Comment == null || Cell.Comment.ToString() == "")
            {
                var analyzerExtension = new tmpAnalyzerExtension(smell.AnalysisType);
                var comments = analyzerExtension.GetSmellMessage(smell);
                if (!string.IsNullOrEmpty(comments))
                {
                    Cell.AddComment(comments);
                }
            }

            Cell.Comment.Visible = true;
        }
    }

    public partial class BBAddIn
    {


        public Ribbon1 theRibbon;
        List<FSharpTransformationRule> AllTransformations = new List<FSharpTransformationRule>();
        public AnalysisController AnalysisController;
        private List<SmellyCell> coloredCells = new List<SmellyCell>();


        private static string RemoveFirstSymbol(string input)
        {
            input = input.Substring(1, input.Length - 1);
            return input;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            theRibbon = new Ribbon1();
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { theRibbon });
            
        }

        public void InitializeBB()
        {
            //initialize smell controls
            theRibbon.selectSmellType.Items.Clear();
            theRibbon.selectSmellType.Enabled = false;

            //initialize transformations
            
            Excel.Worksheet Sheet = GetWorksheetByName("Transformations");

            //find last filled cells
            int Lower = 50;

            for (int i = 1; i <= Lower; i++)
            {
                string From = Sheet.Cells.Item[i, 1].Value;
                if (From != null)
                {
                    string To = Sheet.Cells.Item[i, 2].Value;
                    double prio = Sheet.Cells.Item[i, 3].Value;
                    string Name = Sheet.Cells.Item[i, 4].Value;

                    FSharpTransformationRule S = new FSharpTransformationRule();
                    S.from = S.ParseToTree(From);
                    S.to = S.ParseToTree(To);
                    S.priority = prio;
                    S.Name = Name;

                    AllTransformations.Add(S);

                }
            }


            //order by priority
            TransformationComparer T = new TransformationComparer();

            AllTransformations.Sort(T.Compare);          
        }

 


        public void FindApplicableTransformations()
        {
            Log("FindApplicableTransformations");

            theRibbon.dropDown1.Items.Clear();
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            Excel.Range R = ((Excel.Range)Application.Selection);
            string Formula = R.Formula;

            if (Formula.Length > 0)
            {
                Formula = RemoveFirstSymbol(Formula);

                foreach (FSharpTransformationRule t in AllTransformations)
                {
                    if (t.CanBeAppliedonBool(Formula))
                    {
                        RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                        item.Label = t.Name;
                        theRibbon.dropDown1.Items.Add(item);
                    }
                }
                if (AllTransformations.Count > 0)
                {
                    MakePreview();
                } 
            }     
        }

        private void Log(string LogMessage)
        {
            string currentWorkbookFilePath = this.Application.ActiveWorkbook.Path;
            string LogFileName = "spreadsheets.log";
            string LogFile = System.IO.Path.Combine(currentWorkbookFilePath, LogFileName);
            var file = new System.IO.StreamWriter(LogFile, true);
            file.WriteLine(DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortTimeString() + ", " + LogMessage);
            file.Close();
        }

        public void MakePreview()
        {
            if (theRibbon.dropDown1.Items.Count > 0) //if we have transformations available
            {
                //get the transformation
                FSharpTransformationRule T = AllTransformations.FirstOrDefault(x => x.Name == theRibbon.dropDown1.SelectedItem.Label);

                Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
                Excel.Range R = ((Excel.Range)Application.Selection);

                if (R.Count == 1)
                {
                    string Formula = RemoveFirstSymbol(R.Formula);
                    theRibbon.Preview.Text = T.ApplyOn(Formula);
                }
                else
                {
                    //get first cell for preview
                    Excel.Range Cell1 = R.Cells.Item[0, 0];
                    string Formula = RemoveFirstSymbol(Cell1.Formula);
                    theRibbon.Preview.Text = T.ApplyOn(Formula);
                }
            }
        }

        public void ApplyEverywhere()
        {
            Log("ApplyEverywhere, " + theRibbon.dropDown1.SelectedItem.Label);
            //get the transformation
            FSharpTransformationRule T = AllTransformations.FirstOrDefault(x => x.Name == theRibbon.dropDown1.SelectedItem.Label);

            foreach (var item in Application.Worksheets)
            {
                Excel.Worksheet Worksheet = ((Excel.Worksheet)item);
                Excel.Range R = ((Excel.Range)Worksheet.Cells);

                //find last filled cells
                int Lower = 50;//R.Cells.Find("*", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSearchDirection.xlNext, Type.Missing, Type.Missing, Type.Missing).Row;
                int RightMost = 50;//R.Cells.Find("*", Type.Missing, Excel.XlFindLookIn.xlValues, Type.Missing, Type.Missing, Excel.XlSearchDirection.xlNext, Type.Missing, Type.Missing, Type.Missing).Column;

                for (int i = 1; i <= Lower; i++)
                {
                    for (int j = 1; j <= RightMost; j++)
                    {
                        Excel.Range Cell = R.Cells.Item[i, j];
                        string Formula = Cell.Formula;
                        if (Cell.HasFormula)
                        {
                            Formula = RemoveFirstSymbol(Formula);
                            if (T.CanBeAppliedonBool(Formula))
                            {
                                Cell.Formula = "=" + T.ApplyOn(Formula);
                                Cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                            }
                        }
                    }
                }
            }

        }

        public void ApplyinSheet()
        {
            Log("ApplyinSheet, " + theRibbon.dropDown1.SelectedItem.Label);
            //get the transformation
            FSharpTransformationRule T = AllTransformations.FirstOrDefault(x => x.Name == theRibbon.dropDown1.SelectedItem.Label);

            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            Excel.Range R = ((Excel.Range)activeWorksheet.Cells);

            //find last filled cells
            int Lower = 50;//R.Cells.Find("*", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSearchDirection.xlNext, Type.Missing, Type.Missing, Type.Missing).Row;
            int RightMost = 50;//R.Cells.Find("*", Type.Missing, Excel.XlFindLookIn.xlValues, Type.Missing, Type.Missing, Excel.XlSearchDirection.xlNext, Type.Missing, Type.Missing, Type.Missing).Column;
        
            for (int i = 1; i <= Lower; i++)
            {
                for (int j = 1; j <= RightMost; j++)
                {
                    Excel.Range Cell = R.Cells.Item[i, j];
                    string Formula = Cell.Formula;
                    if (Cell.HasFormula)
                    {
                        Formula = RemoveFirstSymbol(Formula);
                        if (T.CanBeAppliedonBool(Formula))
                        {
                            Cell.Formula = "=" + T.ApplyOn(Formula);
                            Cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                    }
                }    
            }

         }



        public void ApplyinRange()
        {
            if (theRibbon.dropDown1.SelectedItem != null)
            {
                Log("ApplyinRange, " + theRibbon.dropDown1.SelectedItem.Label);
                //get the transformation

                FSharpTransformationRule T = AllTransformations.FirstOrDefault(x => x.Name == theRibbon.dropDown1.SelectedItem.Label);

                Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
                Excel.Range R = ((Excel.Range)Application.Selection);

                foreach (Excel.Range Cell in R.Cells)
                {
                    string Formula = RemoveFirstSymbol(Cell.Formula);
                    Cell.Formula = "=" + T.ApplyOn(Formula);
                }
            }
            else
            {
                Log("ApplyinRange tried with empty dropdown");
            }


            //after applying, we want to empty the preview box, find new rewrites and show them (in dropdown and preview)
            FindApplicableTransformations();
            MakePreview();                                
        }

        public void ColorSmells()
        {
            SpreadsheetInfo.SetLicense("E7OS-D3IG-PM8L-A03O");

            if (!Application.ActiveWorkbook.Saved)
            {
                Application.Dialogs[Excel.XlBuiltInDialog.xlDialogSaveAs].Show();
            }

            if (!Application.ActiveWorkbook.Saved) {
                MessageBox.Show("The workbook must be saved before analysis. Aborting.");
                return;
            }

            AnalysisController = new AnalysisController
            {
                Worker = new BackgroundWorker { WorkerReportsProgress = true },
                AnalysisMaxRows = 10000,
                Filename = Application.ActiveWorkbook.FullName,
                WriteColoredXLSFile = false
            };

            AnalysisController.RunAnalysis(true, false);

            if (!AnalysisController.Spreadsheet.AnalysisSucceeded)
            {
                throw new Exception(AnalysisController.Spreadsheet.ErrorMessage);
            }

            ColorSmellsOfType("");

            LoadSmellTypesSelect();
        }

        public void SelectSmellsOfType()
        {
            ColorSmellsOfType(theRibbon.selectSmellType.SelectedItem.Tag);
        }

        private void ColorSmellsOfType(String type)
        {
            decolorCells();

            List<Smell> smellsOfType;

            if (type == "")
            {
                smellsOfType = AnalysisController.DetectedSmells;
            }
            else
            {
                smellsOfType = AnalysisController.DetectedSmells.Where(x => x.AnalysisType.ToString() == type).ToList();
            }

            foreach (var smell in smellsOfType)
            {
                var analyzerExtension = new tmpAnalyzerExtension(smell.AnalysisType);
                if (analyzerExtension.GetMetricScore(smell.RiskValue) > MetricScore.None) ColorCell(smell);
            }
        }

        private void decolorCells()
        {
            foreach (SmellyCell smellyCell in coloredCells)
            {
                smellyCell.Reset();
            }
            coloredCells.Clear();
        }

        private void ColorCell(Smell smell)
        {
            if (!smell.IsCellBased()) return;

            try
            {
                var cell = (smell.SourceType == RiskSourceType.SiblingClass) ? ((SiblingClass)smell.Source).Cells[0] : (Cell)smell.Source;

                var excelCell = Application.Sheets[cell.Worksheet.Name].Cells[cell.Location.Row + 1, cell.Location.Column + 1];

                var smellyCell = new SmellyCell(excelCell, excelCell.Interior.Pattern, excelCell.Interior.Color);
                smellyCell.Apply(smell);
                coloredCells.Add(smellyCell);
            }
            catch (Exception e)
            {
            }
        }

        public void LoadSmellTypesSelect()
        {
            theRibbon.selectSmellType.Items.Clear();

            foreach (var smellType in AnalysisController.DetectedSmells.Select(x => x.AnalysisType).Distinct())
            {
                tmpAnalyzerExtension analyzerExtension = new tmpAnalyzerExtension(smellType);
                addSelectSmellTypeItem(smellType.ToString(), analyzerExtension.SmellName);
            }

            addSelectSmellTypeItem("", "(all)", true);

            if(theRibbon.selectSmellType.Items.Count > 1) theRibbon.selectSmellType.Enabled = true;
        }

        private void addSelectSmellTypeItem(string id, string option, bool selected = false)
        {
            RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            item.Label = option;
            item.Tag = id;
            theRibbon.selectSmellType.Items.Add(item);
            if(selected) theRibbon.selectSmellType.SelectedItem = item;
        }

        private Excel.Worksheet GetWorksheetByName(string name)
        {
            foreach (Excel.Worksheet worksheet in Application.Worksheets)
            {
                if (worksheet.Name == name)
                {
                    return worksheet;
                }
            }
            throw new ArgumentException();
        }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }

              

        void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            InitializeBB();
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {    
            this.Startup += new EventHandler(ThisAddIn_Startup);
            Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
            Application.SheetSelectionChange += new Excel.AppEvents_SheetSelectionChangeEventHandler(Application_SheetSelectionChange);
}

        void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            theRibbon.Preview.Text = "";
            theRibbon.dropDown1.Items.Clear();
        }










        
        #endregion
    }
}
