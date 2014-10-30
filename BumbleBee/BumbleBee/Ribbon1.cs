﻿#define DEBUG

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn3
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.FindApplicableTransformations();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.ApplyinRange();
        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.MakePreview();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.ApplyEverywhere();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.ApplyinSheet();
        }

    #if DEBUG 
        private void Initialize_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.InitializeBB();
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.ColorSmells();
        }

        private void selectSmellType_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Globals.BBAddIn.SelectSmellsOfType();
        }
    }

    #endif
}
