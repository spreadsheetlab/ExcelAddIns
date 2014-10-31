using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using Infotron.PerfectXL.DataModel;
using GemBox.Spreadsheet;
using AragornAddIn;

namespace AragornAddIn
{
    public class AragornWorkbookClass 
    {


        public Spreadsheet spreadsheet;
        public List<Excel.Worksheet> sheetList=new List<Excel.Worksheet>();
        public Excel.Workbook wkbook;
        


    }
}
