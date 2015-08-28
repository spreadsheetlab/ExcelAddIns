using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using QuickGraph;
using QuickGraph.Serialization;
using QuickGraph.Algorithms;
using QuickGraph.Collections;
using QuickGraph.Contracts;
using QuickGraph.Data;
using QuickGraph.Graphviz;
using QuickGraph.Predicates;
using System.IO;

namespace Polaris
{
    public sealed class FileDotEngine : IDotEngine
    {
        public string Run(QuickGraph.Graphviz.Dot.GraphvizImageType imageType, string dot, string outputFileName)
        {
            string output = outputFileName;
            File.WriteAllText(output, dot);

            // assumes dot.exe is on the path:
            var args = string.Format(@"{0} -Tjpg -O", output);
            System.Diagnostics.Process.Start("C:\\Program Files (x86)\\Graphviz2.38\\bin\\dot.exe", args);
            return output;
        }
    }
}
