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
            

            
        }


        public void TurnOnAragorn()
        {
            //MessageBox.Show("Ready to roll!");
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet); //select active worksheet
            Excel.Range R = ((Excel.Range)Application.Selection); // points to the active selected cell or range

            if (R.Top - 70 <= 0)
            {
                if (R.Left - 140 <= 0)
                {
                    textbox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, R.Left + R.Width, R.Top, 140, 90);

                }
                else
                {
                    textbox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, R.Left - 140, R.Top, 140, 90);

                }
            }
            else
            {
                if (R.Left - 140 <= 0)
                {
                    textbox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, R.Left + R.Width, R.Top - 70, 140, 90);

                }
                else
                {
                    textbox = activeWorksheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, R.Left - 140, R.Top - 70, 140, 90);

                }
            }

            textbox.TextEffect.Text = "Beware! This cell is being used in formulas contained in cells " + R.DirectDependents.get_Address(false);
            textbox.Fill.ForeColor.RGB = 0x87CEEB;

            popupDelay = new System.Timers.Timer(3000);
            popupDelay.Start();
            popupDelay.Elapsed += new ElapsedEventHandler(VanishPopup);
            
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
