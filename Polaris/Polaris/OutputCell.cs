using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FileHelpers;

namespace Polaris
{
    [DelimitedRecord(",")]
    class OutputCell
    {
        public int Id;
        public string WorkbookName;
        public string WorksheetName;
        public string CellAddress;
    }
}
