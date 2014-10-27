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
        
        //Excel.Workbook activeWorkbook;

        //Excel.Sheets sheetCollection;

        //List<Excel.Worksheet> sheetList=new List<Excel.Worksheet>();

        List<AragornWorkbookClass> workbookList = new List<AragornWorkbookClass>();
        int popUpCount = 0;

        Boolean popUpDeleteLock = false;

        Boolean wkBookClosure = false;

        
       
        
        Queue<PopUp> popUpQueue = new Queue<PopUp>();
       // Boolean queueWrite = true;
        Boolean aragornOff = true;
        int aragornTurnedOn = 0;
       
        
        
        //Spreadsheet spreadsheet; // Declare the spreadsheet as a class variable
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e) //executed on startup of excel
        {

            //Globals.Ribbons.Ribbon1.editBox1.Label = "Kutt";
            //Globals.Ribbons.Ribbon1.editBox1.Text = "Kutt";
             //Globals.Ribbons.Ribbon1.editBox1.
            

            //MessageBox.Show("Inside Startup");
            //Excel.Workbook activeWorkbook = ((Excel.Workbook)Application.ActiveWorkbook); //select active workbook
            //activeWorkbook.SheetDeactivate += new Excel.WorkbookEvents_SheetDeactivateEventHandler(activeWorkbook_SheetDeactivate);
            
            
        }

        public void ProcessWorkBook()
        {
            try
            {

                if(workbookList.Exists(w => w.wkbook.Name == Application.ActiveWorkbook.Name))
                {
                    MessageBox.Show("This workbook has been already processed");
                }
                
                else
                {
                    MessageBox.Show("Kindly wait while the workbook is being processed");

                    AragornWorkbookClass workbook = new AragornWorkbookClass();

                    workbook.wkbook = ((Excel.Workbook)Application.ActiveWorkbook); //select active workbook

                    AnalysisController c = new AnalysisController();
                    c.AnalysisMaxRows = 10000;
                    //spreadsheet = new Spreadsheet();
                    SpreadsheetInfo.SetLicense("E7OS-D3IG-PM8L-A03O");
                    workbook.spreadsheet = c.OpenSpreadsheet(Application.ActiveWorkbook.FullName, analyzeAllSiblings: false, precedentsForAllSiblings: true);



                    CreateWorkbookEventHandlers(workbook);

               

                    CreateSheetEventHandlers(workbook);

                    workbookList.Add(workbook);

                    MessageBox.Show("AraSENSE is ready for activation");

                }
                


                //activeWorkbook.SheetDeactivate += new Excel.WorkbookEvents_SheetDeactivateEventHandler(activeWorkbook_SheetDeactivate);
            } 
            
            catch (Exception  e)
            {
                MessageBox.Show("Error!\nPlease load a proper Excel file with .xls or .xlsx extension first in order to process\nError Message: " + e);
            }

                        
        }

        private void CreateWorkbookEventHandlers(AragornWorkbookClass workbook)
        {
            workbook.wkbook.BeforeClose += new Excel.WorkbookEvents_BeforeCloseEventHandler(activeWorkbook_BeforeClose);

            //activeWorkbook.SheetChange += new Excel.WorkbookEvents_SheetChangeEventHandler(activeWorkbook_SheetChange);
            workbook.wkbook.AfterSave += new Excel.WorkbookEvents_AfterSaveEventHandler(activeWorkbook_AfterSave);
        }

        private void CreateSheetEventHandlers(AragornWorkbookClass workbook)
        {


            Excel.Sheets sheetCollection = workbook.wkbook.Sheets;

            //MessageBox.Show("Number of sheets " + activeWorkbook.Sheets.Count);
            if (workbook.sheetList.Count != 0)
            { workbook.sheetList.Clear(); }

            for (int i = 1; i <= sheetCollection.Count; i++)
            {
                workbook.sheetList.Add(sheetCollection[i]);

                workbook.sheetList[i - 1].SelectionChange += new Excel.DocEvents_SelectionChangeEventHandler(activeWorksheet1_SelectionChange); //the event handler for on change of cell event





            }
        }

        void activeWorkbook_AfterSave(bool Success)
        {
            //MessageBox.Show("Inside Workbook Saved");
            ReProcessWorkbook();
        }

        

        private void ReProcessWorkbook()
        {
            
            if(wkBookClosure==false)
            {
                MessageBox.Show("You have saved modifications to the workbook. Kindly wait till it is reprocessed.");
                //activeWorkbook.Save();
                AnalysisController c = new AnalysisController();
                c.AnalysisMaxRows = 10000;
                Spreadsheet spreadsheet = new Spreadsheet();
                SpreadsheetInfo.SetLicense("E7OS-D3IG-PM8L-A03O");
                spreadsheet = c.OpenSpreadsheet(Application.ActiveWorkbook.FullName, analyzeAllSiblings: false, precedentsForAllSiblings: true);

                AragornWorkbookClass workbook= workbookList.Find(w => w.spreadsheet.Filename == spreadsheet.Filename);
                workbook.spreadsheet = spreadsheet;
                workbook.wkbook = ((Excel.Workbook)Application.ActiveWorkbook);


                DestroySheetEventHandlers(workbook);
                CreateSheetEventHandlers(workbook);

                MessageBox.Show("AraSENSE is ready again.");
            }
            
            
        }

        private void DestroySheetEventHandlers(AragornWorkbookClass workbook)
        {

            //MessageBox.Show("Number of sheets " + sheetList.Count);
            for (int i = 0; i < workbook.sheetList.Count; i++)
            {


                workbook.sheetList[i].SelectionChange -= activeWorksheet1_SelectionChange;





            }
        }

        void activeWorkbook_BeforeClose(ref bool Cancel)
        {
            
            MessageBox.Show("You have chosen to close this workbook. Due to Aragorn all changes will be saved.");
            WorkBookCloseCleanUp();
        }

        

        private void WorkBookCloseCleanUp()
        {
            //MessageBox.Show("Inside sheet dd");



            while (popUpDeleteLock == true) ;
            popUpDeleteLock = true;
            while(popUpQueue.Count!=0)
            {

                
                PopUp popUp = popUpQueue.Dequeue();
                popUpCount--;
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

              
            }
            popUpDeleteLock = false;
            
            wkBookClosure = true;
            Excel.Workbook activeWorkbook = ((Excel.Workbook)Application.ActiveWorkbook);
            activeWorkbook.Save();
            AragornWorkbookClass workbook = workbookList.Find(w => w.wkbook.Name == activeWorkbook.Name);
            workbookList.Remove(workbook);
            wkBookClosure = false;

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
            Globals.Ribbons.Ribbon1.label1.Label = ""; 
            MessageBox.Show("AraSENSE is de-activated");
        }




        private void CellChangeEvent(Excel.Range Target)
        {

            Boolean colonFlag = false;
            //Boolean newWorksheet = false;
            int dependentsCount = 0;
            try
            {
                if (aragornOff == false)
                {
                    PopUp popUp = new PopUp();
                    popUp.popUpText = "";

                    //if (Target.get_Value() != null) //checking for non-empty cell
                    

                        Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet); //select active worksheet
                        //MessageBox.Show(Target.Address);
                        String cellAddress = String.Join("", Target.Address.Split('$'));
                       // MessageBox.Show(cellAddress);

                        Excel.Workbook activeWorkbook = ((Excel.Workbook)Application.ActiveWorkbook);

                        Spreadsheet spreadsheet = workbookList.Find(w => w.wkbook.Name == activeWorkbook.Name).spreadsheet;

                        Cell cell = spreadsheet.GetWorksheet(activeWorksheet.Name).GetCell(cellAddress);
                        if(cell==null)
                        {
                            
                            
                        
                            Globals.Ribbons.Ribbon1.label1.Label = "No Dependents of selected cell detected. If you have modified this spreadsheet please save in order to see the changes relfected"; 
                        
                            return;
                        }
                        //MessageBox.Show("Cell formula from infotron core  "+cell.Formula);//.Location.ToString());
                        List<Cell> dependents = new List<Cell>();
                        dependents = cell.GetDependents();
                        dependentsCount = dependents.Count;
                        //MessageBox.Show("dependentsCount  " + dependentsCount);

                        //Boolean workSheetFlag = false;

                        for (int i = 0; i < dependents.Count; i++) // Loop through List with for
                        {
                            //MessageBox.Show("Iterating List: " + dependents[i].Worksheet.Name);

                            Location loc2 = dependents[i].Location;
                            if (i != 0)
                            {
                                Location loc1 = dependents[i - 1].Location;
                                String str1 = loc1.ToString();
                                String str2 = loc2.ToString();
                                if (dependents[i].Worksheet.Name != dependents[i - 1].Worksheet.Name)
                                {

                                    //popUp.popUpText = popUp.popUpText + "\n<Sheet " + dependents[i].Worksheet.Name + ">! ";
                                    
                                    if (colonFlag == true)
                                    {
                                        
                                        popUp.popUpText = popUp.popUpText + str1 + " ";
                                        colonFlag = false;
                                    }
                                    popUp.popUpText = popUp.popUpText + "\n<Sheet " + dependents[i].Worksheet.Name + ">! " + str2 + " ";


                                }
                                
                                else
                                {
                                    if ((loc1.Row == loc2.Row)&&((loc2.Column-loc1.Column)==1))
                                    {
                                        if (colonFlag == false)
                                        {
                                            popUp.popUpText = popUp.popUpText + ":";
                                            colonFlag = true;
                                        }

                                    }
                                    else
                                    {

                                        if (colonFlag == true)
                                        {

                                            popUp.popUpText = popUp.popUpText + str1 + " ";
                                            colonFlag = false;
                                        }
                                        popUp.popUpText = popUp.popUpText + str2 + " ";


                                    }
                                }
                                

                                
                            }
                            

                            else
                            {
                                String str = loc2.ToString();
                                if (dependents[i].Worksheet.Name != activeWorksheet.Name)
                                { popUp.popUpText = popUp.popUpText + "\n<Sheet " + dependents[i].Worksheet.Name + ">! "; }
                                popUp.popUpText = popUp.popUpText + str + " ";
                                
                            }

                            
                            //MessageBox.Show("Inside list: " + str);
                            

                        }

                        // MessageBox.Show("Iterating List: " + popUp.popUpText);

                        
                        
                        
                        if (popUp.popUpText != "")
                        {


                            //if (Target.Top - 70 <= 0)
                            //{
                            //    if (Target.Left - 140 <= 0)
                            //    {
                            //        popUp.textBox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left + Target.Width, Target.Top, 140, 130);

                            //    }
                            //    else
                            //    {
                            //        popUp.textBox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left - 140, Target.Top, 140, 130);

                            //    }
                            //}
                            //else
                            //{
                            //    if (Target.Left - 140 <= 0)
                            //    {
                            //        popUp.textBox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left + Target.Width, Target.Top - 70, 140, 130);

                            //    }
                            //    else
                            //    {
                            //        popUp.textBox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left - 140, Target.Top - 70, 140, 130);

                            //    }


                           

                            //}
                            popUp.textBox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 140, 130);


                            popUp.textBox.TextFrame2.TextRange.Text = "Dependents\n=============\n" + "No. of Dependents: " + dependentsCount + "\n\n" + popUp.popUpText;//+ ;

                            /**********/
                            //Globals.Ribbons.Ribbon1.editBox1.Label = "Beware! Dependents sensed >>\n" + popUp.popUpText;

                            Globals.Ribbons.Ribbon1.label1.Label = "No. of Dependents: " + dependentsCount+"==>>"+popUp.popUpText; 
                            /*********/

                            popUp.textBox.TextFrame2.WordWrap = (Office.MsoTriState) 1;

                            popUp.textBox.TextFrame2.AutoSize = (Office.MsoAutoSize) 1; 

                            popUp.textBox.Fill.ForeColor.RGB = 0x87CEEB;

                            

                            popUp.popupDelay = new System.Timers.Timer(3000);
                            popUp.popupDelay.Start();


                            popUpQueue.Enqueue(popUp);
                            popUpCount++;


                            popUp.popupDelay.Elapsed += new ElapsedEventHandler(popupDelay_Elapsed); //+= new ElapsedEventHandler(VanishPopup);
                            //throw new NotImplementedException();
                        }
                        else
                        {
                            Globals.Ribbons.Ribbon1.label1.Label = "No Dependents of selected cell detected. If you have modified this spreadsheet please save in order to see the changes relfected"; 
                        }
                       
                    
                }
            }

            catch (Exception e)
            {
               MessageBox.Show("Error!\nPlease try selecting another 'single' cell please\nError Message: " + e);
            }
            
        }

        void popupDelay_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                while (popUpDeleteLock == true) ;
                if (popUpQueue.Count != 0)
                {


                    popUpDeleteLock = true;
                    PopUp popUp = new PopUp();
                    popUp = popUpQueue.Dequeue();
                    popUpCount--;
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



                    popUpDeleteLock = false;

                }
            

            }catch (NullReferenceException es)
            { 
                MessageBox.Show("Fatal Error! (I know this is crappy but we have to work on it. Really sorry)");
                
            
            }
           
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
