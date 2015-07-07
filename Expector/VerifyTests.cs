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
            TestCheck t = new TestCheck(f);
            

            int Passing = P.GetPassingBranch(f.original);

            string Condition = P.GetCondition(f.original);

            List<String> ConditionList = P.Split(Condition);

            string Text;

            if (ConditionList.Count == 1) //just 1 item in the condition, like IF(A4,"OK","NOT OK")
            {
                Text = Condition + " should ";

                if (P.GetPassingBranch(f.original) == 0) //if the first branch is passing, the test condition should be true, else false
                //the senmatic of IF(A4) is that it is false if A4 = 0, true in all other cases.
                {
                    Text += "not be 0"; //it should be true, so non-zero
                    t.shouldbe = true;
                }
                else
                {
                    Text += "be 0";
                    t.shouldbe = false;
                }
            }
            else
            {
                if (ConditionList.Count == 2) //a function as condition 1, like IF(ISBLANK(A4),"OK","NOT OK")
                {
                    Text = Condition + " should ";
                    if (P.GetPassingBranch(f.original) == 0) //if the first branch is passing, the test condition should be true, else false                                    
                    {
                        Text += "hold"; //it should be true, so non-zero
                        t.shouldbe = true;
                    }
                    else
                    {
                        Text += "not hold";
                        t.shouldbe = false;
                    }
                }

                else //the condition is a 'real' condition i.e. IF(A5 > 5, "OK", "ERROR")
                {
                    Text = ConditionList[0];

                    if (P.GetPassingBranch(f.original) == 0) //if the first branch is passing, the test codintion should be true, else false
                    {
                        Text += " should be " + ConditionList[1] + ConditionList[2]; //this adds the > 5 part
                        t.shouldbe = true;
                    }
                    else
                    {
                        Text += " should be " + Invert(ConditionList[1]) + ConditionList[2]; //this adds the <= 5 part
                        t.shouldbe = false;
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
            l.Text = Formula.worksheet + "!" +Formula.location + " tests "+ t.outputText;
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
            List<testFormula> formulas = new List<testFormula>();

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

            instanceofExpector.TestFormulas = formulas;
            instanceofExpector.SaveTests();

            this.Close();
        }
    }


}
