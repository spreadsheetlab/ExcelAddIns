using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FileHelpers;

namespace Columba2
{
    [DelimitedRecord(",")]
    class CSV_Charts
    {
        public string WorkbookName;
        public string WorksheetName;
        public string ChartName;
        public string ChartTitle;
        public string ChartType;
    }
}
