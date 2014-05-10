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

namespace AragornAddIn
{
    public partial class ThisAddIn
    {

        Excel.Shape textbox; // Declare the textbox as a class variable
        System.Timers.Timer popupDelay; //Declare the delay for lasting the popups
        Spreadsheet spreadsheet; // Declare the spreadsheet as a class variable
        String popUp="" ; // the string to contain celle references shown in the popup
        private void ThisAddIn_Startup(object sender, System.EventArgs e) //executed on startup of excel, analyzes whole sheet
        {

            //Boolean analyzeAllSiblings = true;
            //Controller c = new Controller();
            //spreadsheet = new Spreadsheet();

            //String fileName = @"C:\Copy of 66.xlsx";
            //if (String.Equals(fileName, Application.ActiveWorkbook.FullName))
            //{ MessageBox.Show("EQUAL"); }

            //MessageBox.Show(fileName + "\n" + Application.ActiveWorkbook.FullName+"Q");

            //spreadsheet = c.OpenSpreadsheet(fileName, analyzeAllSiblings);//(@"C:\Copy of 66.xlsx", analyzeAllSiblings);
            

            

            
        }


        public void TurnOnAragorn() //executed on ON button click
        {

            Boolean analyzeAllSiblings = true;
            Controller c = new Controller();
            spreadsheet = new Spreadsheet();
           
            spreadsheet = c.OpenSpreadsheet(Application.ActiveWorkbook.FullName, analyzeAllSiblings);
            
            
            
            MessageBox.Show("AraSENSE is activated");
            

           

            Excel.Worksheet activeWorksheet1 = ((Excel.Worksheet)Application.ActiveSheet); //select active worksheet
            
            
            
            activeWorksheet1.SelectionChange += new  Excel.DocEvents_SelectionChangeEventHandler(activeWorksheet1_SelectionChange); //the event handler for on change of cell event

           
            
        }



        void activeWorksheet1_SelectionChange(Excel.Range Target) //the method to handle the change of cell event, shows the popup
        {

            if (Target.get_Value()!=null)
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
                           
                                popUp = popUp +  "\n<Sheet "+dependents[i].Worksheet.Name + ">: ";
                        }
                    }
                    Location loc = dependents[i].Location;
                    String str = loc.ToString();
                    //MessageBox.Show("Inside list: " + str);
                    popUp = popUp + str + " ";

                }

                if (popUp != "")
                {
                    if (Target.Top - 70 <= 0)
                    {
                        if (Target.Left - 140 <= 0)
                        {
                            textbox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left + Target.Width, Target.Top, 140, 130);

                        }
                        else
                        {
                            textbox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left - 140, Target.Top, 140, 130);

                        }
                    }
                    else
                    {
                        if (Target.Left - 140 <= 0)
                        {
                            textbox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left + Target.Width, Target.Top - 70, 140, 130);

                        }
                        else
                        {
                            textbox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left - 140, Target.Top - 70, 140, 130);

                        }
                    }

                    textbox.TextEffect.Text = "Beware! Dependents sensed >>\n" + popUp;//+ ;
                    textbox.Fill.ForeColor.RGB = 0x87CEEB;

                    popupDelay = new System.Timers.Timer(3000);
                    popupDelay.Start();
                    popupDelay.Elapsed += new ElapsedEventHandler(VanishPopup);
                    //throw new NotImplementedException();
                }
            }
        }


        private void VanishPopup(object source, ElapsedEventArgs e)
        {

            textbox.Delete();
            popUp = "";
            popupDelay.Stop();
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
