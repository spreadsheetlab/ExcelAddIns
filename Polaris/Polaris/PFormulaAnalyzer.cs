using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XLParser;
using Irony.Parsing;

namespace Polaris
{
    class PFormulaAnalyzer : FormulaAnalyzer
    {

        public PFormulaAnalyzer(string formula) : base(formula)
        { }
        public IEnumerable<string> BuiltinFunctions()
        {
            return AllNodes
                .Where(node => node.IsBuiltinFunction())
                .Select(ExcelFormulaParser.GetFunction);
        }

    }
}
