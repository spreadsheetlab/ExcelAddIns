using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Runtime.InteropServices;
using FileHelpers;

namespace Columba2
{
    class Scanner
    {
        private List<CSV_Charts> foundCharts;
        public void ExecuteScan()
        {
            foundCharts = new List<CSV_Charts>();
            string folderToScan = GetFolderToScan();
            // Cancel scan if no folder was selected
            if (folderToScan == "")
            {
                MessageBox.Show("No folder was selected, the scan was aborted");
                return;
            }
            // Get Excel files to scan
            string[] excelFiles = GetExcelFiles(folderToScan);
            // Check if any files are returned
            if (excelFiles.Length == 0)
            {
                MessageBox.Show("No Excel files were found, the scan was aborted");
                return;
            }

            // Connect to Excel Application
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            xlApp.ScreenUpdating = false;

            // Execute scan for every file
            foreach (string excelFile in excelFiles)
            {
                // Open workbook
                xlApp.DisplayAlerts = false;
                Excel.Workbooks xlWorkbooks = xlApp.Workbooks;
                Excel.Workbook wkb = xlWorkbooks.Open(Filename: excelFile, ReadOnly: true, UpdateLinks: 2);
                xlApp.DisplayAlerts = false;

                // Check if there are any charts in the workbook, these are the so called chart sheets
                Excel.Sheets charts = wkb.Charts;
                if (charts.Count != 0)
                {
                    for (int i = 1; i <= charts.Count; i++) // chart index starts at 1
                    {
                        AddChart(charts[i], wkb);
                    }
                }

                // Charts can also be embedded objects on the worksheet. Therefore we have to check every worksheet for charts
                for (int i = 1; i <= wkb.Worksheets.Count; i++) // worksheet index starts at 1
                {
                    Excel.Worksheet wks = wkb.Worksheets[i];
                    Excel.ChartObjects chartObjects = wks.ChartObjects();
                    if (chartObjects.Count != 0)
                    {
                        for (int j = 1; j <= chartObjects.Count; j++)
                        {
                            AddChart(chartObjects.Item(j).Chart, wkb, wks);
                        }
                    }
                    Marshal.FinalReleaseComObject(chartObjects);
                    Marshal.FinalReleaseComObject(wks);
                }

                // Close workbook and cleanup
                wkb.Close(SaveChanges: false);
                Marshal.FinalReleaseComObject(charts);
                Marshal.FinalReleaseComObject(wkb);
                Marshal.FinalReleaseComObject(xlWorkbooks);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            GenerateCSV();
            xlApp.ScreenUpdating = true;
        }
        private void GenerateCSV()
        {
            var engine = new FileHelperEngine<CSV_Charts>();
            engine.AppendToFile("charts.csv", foundCharts);
        }
        private void AddChart(Excel.Chart chart, Excel.Workbook wkb, Excel.Worksheet wks = null)
        {
            CSV_Charts csvChart = new CSV_Charts();
            csvChart.WorkbookName = wkb.Name;
            // If wks is null the chart is a worksheet itself and the name of the chart is the sheetname
            if (wks == null)
            {
                csvChart.WorksheetName = chart.Name; 
            }
            else
            {
                csvChart.WorksheetName = wks.Name;
            }
            if (chart.HasTitle) csvChart.ChartTitle = chart.ChartTitle.Text;
            csvChart.ChartName = chart.Name;
            csvChart.ChartType = chart.ChartType.ToString();

            foundCharts.Add(csvChart);
        }
        private string GetFolderToScan()
        {
            // Display select folder dialog and return selected folder
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.ShowDialog();
            return fbd.SelectedPath;
        }
        private string[] GetExcelFiles(string folderToScan)
        {
            // Get list of all Excel files (xls and xlsx)
            return Directory.GetFiles(folderToScan, "*.*")
                .Where(s => s.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) || 
                    s.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)).ToArray();
        }
    }
}
