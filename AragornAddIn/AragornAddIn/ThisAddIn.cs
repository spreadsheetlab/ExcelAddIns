using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Nl.Infotron.Analyzer.DataModel;
using Nl.Infotron.Analyzer;
using System.Windows.Forms;
using System.Drawing;
using System.Timers;

namespace AragornAddIn
{
    public partial class ThisAddIn
    {

        Excel.Shape textbox;
        System.Timers.Timer popupDelay;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //IAnalyzer a = new DefaultAnalyzer();
            //Excel.Worksheet activeWorksheet1 = ((Excel.Worksheet)Application.ActiveSheet); //select active worksheet
            

            
        }


        public void TurnOnAragorn()
        {
            //MessageBox.Show("Ready to roll!");
            Excel.Worksheet activeWorksheet1 = ((Excel.Worksheet)Application.ActiveSheet); //select active worksheet
            //Excel.Range R = ((Excel.Range)Application.Selection); // points to the active selected cell or range

            activeWorksheet1.SelectionChange += new  Excel.DocEvents_SelectionChangeEventHandler(activeWorksheet1_SelectionChange);

           
            
        }



        void activeWorksheet1_SelectionChange(Excel.Range Target)
        {

            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet); //select active worksheet
            //Excel.Range R = ((Excel.Range)Application.Selection); // points to the active selected cell or range
            MessageBox.Show(" :O :O");
            if (Target.Top - 70 <= 0)
            {
                if (Target.Left - 140 <= 0)
                {
                    textbox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left + Target.Width, Target.Top, 140, 90);

                }
                else
                {
                    textbox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left - 140, Target.Top, 140, 90);

                }
            }
            else
            {
                if (Target.Left - 140 <= 0)
                {
                    textbox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left + Target.Width, Target.Top - 70, 140, 90);

                }
                else
                {
                    textbox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, Target.Left - 140, Target.Top - 70, 140, 90);

                }
            }

            textbox.TextEffect.Text = "Beware! This cell is being used in formulas contained in cells ";//+ Target.DirectDependents.get_Address(false);
            textbox.Fill.ForeColor.RGB = 0x87CEEB;

            popupDelay = new System.Timers.Timer(3000);
            popupDelay.Start();
            popupDelay.Elapsed += new ElapsedEventHandler(VanishPopup);
            //throw new NotImplementedException();
        }


        private void VanishPopup(object source, ElapsedEventArgs e)
        {

            textbox.Delete();
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
