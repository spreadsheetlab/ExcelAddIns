using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Infotron.Parsing;

namespace Expector
{


    public partial class VerifyTests : Form
    {
        public List<TestCheck> TestsChecked = new List<TestCheck>();

        Expector instanceofExpector;

        public VerifyTests(Expector t)
        {
            instanceofExpector = t;
            InitializeComponent();
        }

        public TestCheck ProcessConditions(testFormula f, ExcelFormulaParser P)
        {

            int PassingBranch = P.GetPassingBranch(f.original);
            string Condition = P.GetCondition(f.original);
            List<String> ConditionList = P.Split(Condition);

            string Text;

            TestCheck t = new TestCheck(f);
            t.shouldbe = (PassingBranch == 0); //if the first branch is passing, the condition should be true, else false

            if (ConditionList.Count == 1) //just 1 item in the condition, like IF(A4,"OK","NOT OK")
            {
                Text = Condition + " should ";

                if (PassingBranch == 0) //if the first branch is passing, the test condition should be true, else false
                //the senmatic of IF(A4) is that it is false if A4 = 0, true in all other cases.
                {
                    Text += "not be 0"; //it should be true, so non-zero
                }
                else
                {
                    Text += "be 0";
                }
            }
            else
            {
                if (ConditionList.Count == 2) //if the split found 2 items, the condition is a function, like IF(ISBLANK(A4),"OK","NOT OK")
                {
                    Text = Condition + " should ";

                    if (PassingBranch == 0) //if the first branch is passing, the test condition should be true, else false                                    
                    {
                        Text += "be true"; //it should be true, so non-zero
                    }
                    else
                    {
                        Text += "not be true";
                    }
                }

                else //the condition is a 'real' condition i.e. IF(A5 > 5, "OK", "ERROR")
                {
                    Text = ConditionList[0];

                    if (PassingBranch == 0) //if the first branch is passing, the test codintion should be true, else false
                    {
                        Text += " should be " + ConditionList[1].Replace("=","") + ConditionList[2]; //this adds the > 5 part
                    }
                    else
                    {
                        Text += " should be " + Invert(ConditionList[1]).Replace("=", "") + ConditionList[2]; //this adds the < 5 part, we remove '=' because "x should be =0" looks strange
                    }
                }
            }

            t.outputText = Text;
            return t;
        }


        private string Invert(string p)
        {
            if (p == ">") return "<=";
            if (p == "<") return ">=";
            if (p == ">=") return "<";
            if (p == "<=") return ">";
            if (p == "<>") return "=";
            if (p == "=") return "<>";

            return "not" + p;
        }

        internal void PrintTest(testFormula Formula)
        {
            TestCheck t =  ProcessConditions(Formula, new ExcelFormulaParser());

            int NTestsPrinted = TestsChecked.Count;

            int height = 41 + 23 * NTestsPrinted;

            Label l = new Label();
            l.Location = new Point(12, height);
            l.Text = "The cell on " + Formula.worksheet + "!" +Formula.location + " expresses that "+ t.outputText;
            l.AutoSize = true;
            this.Controls.Add(l);

            CheckBox C = new CheckBox();
            C.Checked = true;
            C.Location = new Point(450, height);
            this.Controls.Add(C);

            t.AddCheckbox(C);

            TestsChecked.Add(t);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<testFormula> formulas = instanceofExpector.testFormulas;

            foreach (var c in TestsChecked)
	        {
                int i = 0;
		        if (c.checkbox.Checked)
	            {
                    formulas.Add(c);
                    i++;
	            }

                //transform the original formula to a condition that always should be true
                ExcelFormulaParser P = new ExcelFormulaParser();

                string formula = P.GetCondition(c.original, c.worksheet);

                if (c.shouldbe == false)
                {
                    c.condition = "NOT(" + formula + ")";
                }
                else
                {
                    c.condition = formula;
                }
	        }

            instanceofExpector.testFormulas = formulas;
            instanceofExpector.SaveTests();

            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            foreach (TestCheck c in TestsChecked)
            {
                c.checkbox.Checked = false;
            }
        }
    }
    public class TestCheck : testFormula
    {
        public CheckBox checkbox;
        public bool shouldbe; // we only need to save this for non-transformed formulas, because we make them as always should be true
        public string outputText;

        public TestCheck(testFormula Formula)
        {
            original = Formula.original;
            worksheet = Formula.worksheet;
            condition = Formula.condition;
            location = Formula.location;
        }

        public void AddCheckbox(CheckBox Checkbox)
        {
            checkbox = Checkbox;
        }

    }

}
