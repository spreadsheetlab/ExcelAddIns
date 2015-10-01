using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Polaris
{
    class PrecedentCell : IDisposable
    {
        private Excel.Range thisCell;
        private Excel.Range dependentCell;
        private int level;
        private string id;
        public PrecedentCell(Excel.Range c, int l, Excel.Range dependent)
        {
            thisCell = c;
            dependentCell = dependent;
            level = l;
            Excel.Worksheet wks = c.Parent;
            id = wks.Name + "!" + c.Address;
            Marshal.FinalReleaseComObject(wks);
        }
        public Excel.Range ThisCell
        {
            get
            {
                return thisCell;
            }
        }
        public Excel.Range DependentCell
        {
            get
            {
                return dependentCell;
            }
        }
        public int Level
        {
            get
            {
                return level;
            }
            set
            {
                level = value;
            }
        }
        public string ID
        {
            get
            {
                return id;
            }
        }
        public void Dispose()
        {
            Marshal.FinalReleaseComObject(thisCell);
            Marshal.FinalReleaseComObject(dependentCell);
        }
    }
}
