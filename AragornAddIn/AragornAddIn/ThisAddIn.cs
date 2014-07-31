using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Drawing;
using System.Timers;
using Infotron.PerfectXL;
using Infotron.PerfectXL.DataModel;
using GemBox.Spreadsheet;
using Infotron.Converter;
using Infotron.Util;
using Infotron.PerfectXL.SmellAnalyzer;
using AragornAddIn;

namespace AragornAddIn
{
    public partial class ThisAddIn
    {
        
        Excel._Workbook activeWorkbook;

        Excel.Sheets sheetCollection;

        List<Excel.Worksheet> sheetList=new List<Excel.Worksheet>();

       
        
        Queue<PopUp> popUpQueue = new Queue<PopUp>();
       // Boolean queueWrite = true;
        Boolean aragornOff = true;
        int aragornTurnedOn = 0;
       
        
        //Excel.Shape textbox; // Declare the textbox as a class variable
        //System.Timers.Timer popupDelay; //Declare the delay for lasting the popups
        Spreadsheet spreadsheet; // Declare the spreadsheet as a class variable
        //String popUpText="" ; // the string to contain celle references shown in the popup
        private void ThisAddIn_Startup(object sender, System.EventArgs e) //executed on startup of excel
        {

           
            //MessageBox.Show("Inside Startup");
            //Excel.Workbook activeWorkbook = ((Excel.Workbook)Application.ActiveWorkbook); //select active workbook
            //activeWorkbook.SheetDeactivate += new Excel.WorkbookEvents_SheetDeactivateEventHandler(activeWorkbook_SheetDeactivate);
            
            
        }

        public void ProcessWorkBook()
        {
            
            AnalysisController c = new AnalysisController();
            spreadsheet = new Spreadsheet();

            SpreadsheetInfo.SetLicense("E7OS-D3IG-PM8L-A03O");

            spreadsheet = c.OpenSpreadsheet(Application.ActiveWorkbook.FullName, analyzeAllSiblings: false, precedentsForAllSiblings: true);
            MessageBox.Show("AraSENSE is ready for activation");

            activeWorkbook = ((Excel.Workbook)Application.ActiveWorkbook); //select active workbook
            sheetCollection = activeWorkbook.Sheets;

            MessageBox.Show("Number of sheets " + activeWorkbook.Sheets.Count);

            for (int i = 1; i < sheetCollection.Count; i++)
            {
                sheetList.Add(sheetCollection[i]);

                sheetCollection[i].SelectionChange += new Excel.DocEvents_SelectionChangeEventHandler(activeWorksheet1_SelectionChange); //the event handler for on change of cell event
               
            }


                //activeWorkbook.SheetDeactivate += new Excel.WorkbookEvents_SheetDeactivateEventHandler(activeWorkbook_SheetDeactivate);

                        
        }

        //void activeWorkbook_SheetDeactivate(object Sh)
        //{

        //    SheetChangeEvent();
            
        //}

        //private void SheetChangeEvent()
        //{

        //    MessageBox.Show("Sheet Changed");

        //    if (aragornOff == false)
        //    {
        //        activeWorksheet1 = ((Excel.Worksheet)Application.ActiveSheet); //select active worksheet
        //        activeWorksheet1.SelectionChange += new Excel.DocEvents_SelectionChangeEventHandler(activeWorksheet1_SelectionChange); //the event handler for on change of cell event

        //    }

        //    else
        //    { aragornTurnedOn = 0; }

        //}

        void activeWorksheet1_SelectionChange(Excel.Range Target) //the method to handle the change of cell event, shows the popup
        {
            CellChangeEvent(Target);
        }

        

               

        public void TurnOnAragorn() //executed on ON button click
        {

           


            aragornOff = false;
            MessageBox.Show("AraSENSE is activated");


            //if (aragornTurnedOn == 0)
            //{

            //    activeWorksheet1 = ((Excel.Worksheet)Application.ActiveSheet); //select active worksheet
            //    activeWorksheet1.SelectionChange += new Excel.DocEvents_SelectionChangeEventHandler(activeWorksheet1_SelectionChange); //the event handler for on change of cell event

            //}

            aragornTurnedOn++; 
            
        }


        public void TurnOffAragorn()
        {
            aragornOff = true;
            //aragornTurnedOn = 0;
            MessageBox.Show("AraSENSE is de-activated");
        }




        private void CellChangeEvent(Excel.Range Target)
        {


            if (aragornOff == false)
            {
                PopUp popUp = new PopUp();
                popUp.popUpText = "";

                if (Target.get_Value() != null) //checking for non-empty cell
                {
                    Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet); //select active worksheet
                    //MessageBox.Show(Target.Address);
                    String cellAddress = String.Join("", Target.Address.Split('$'));
                    //MessageBox.Show(cellAddress);
                    Cell cell = spreadsheet.GetWorksheet(activeWorksheet.Name).GetCell(cellAddress);
                    //MessageBox.Show("Cell formula from infotron core  "+cell.Formula);//.Location.ToString());
                    List<Cell> dependents = cell.GetDependents();

                    //Boolean workSheetFlag = false;

                    for (int i = 0; i < dependents.Count; i++) // Loop through List with for
                    {
                        //MessageBox.Show("Iterating List: " + dependents[i].Worksheet.Name);
                        if (i != 0)
                        {
                            if (dependents[i].Worksheet.Name != dependents[i - 1].Worksheet.Name)
                            {

                                popUp.popUpText = popUp.popUpText + "\n<Sheet " + dependents[i].Worksheet.Name + ">: ";
                            }
                        }
                        Location loc = dependents[i].Location;
                        String str = loc.ToString();
                        //MessageBox.Show("Inside list: " + str);
                        popUp.popUpText = popUp.popUpText + str + " ";

                    }

                   // MessageBox.Show("Iterating List: " + popUp.popUpText);
                    
                    if (popUp.popUpText != "")
                    {


                        if (Target.Top - 70 <= 0)
                        {
                            if (Target.Left - 140 <= 0)
                            {
                                popUp.textBox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left + Target.Width, Target.Top, 140, 130);

                            }
                            else
                            {
                                popUp.textBox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left - 140, Target.Top, 140, 130);

                            }
                        }
                        else
                        {
                            if (Target.Left - 140 <= 0)
                            {
                                popUp.textBox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left + Target.Width, Target.Top - 70, 140, 130);

                            }
                            else
                            {
                                popUp.textBox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left - 140, Target.Top - 70, 140, 130);

                            }
                        }

                        popUp.textBox.TextEffect.Text = "Beware! Dependents sensed >>\n" + popUp.popUpText;//+ ;
                        popUp.textBox.Fill.ForeColor.RGB = 0x87CEEB;

                        popUp.popupDelay = new System.Timers.Timer(3000);
                        popUp.popupDelay.Start();


                        popUpQueue.Enqueue(popUp);


                        popUp.popupDelay.Elapsed += new ElapsedEventHandler(popupDelay_Elapsed); //+= new ElapsedEventHandler(VanishPopup);
                        //throw new NotImplementedException();
                    }
                }
            }
            
        }

        void popupDelay_Elapsed(object sender, ElapsedEventArgs e)
        {
            PopUp popUp = popUpQueue.Dequeue();
            popUp.popupDelay.Stop();
            Boolean deleteFailed = false;
            Boolean userLock = false;
            do
            {
                try
                {
                    do
                    {
                        try
                        {
                            popUp.textBox.Cut();
                            popUp.popUpText = "";

                            deleteFailed = false;
                            userLock = false;
                        }
                        catch (System.UnauthorizedAccessException ex1)
                        {
                            userLock = true;
                        }

                    } while (userLock);



                }

                catch (System.Runtime.InteropServices.COMException ex)
                {
                    deleteFailed = true;
                }
            } while (deleteFailed);

            //throw new NotImplementedException();
        }


     


       

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
