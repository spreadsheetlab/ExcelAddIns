using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Expector
{
    public class TestCheck : testFormula
    {
        public CheckBox checkbox;

        public TestCheck(testFormula Formula, CheckBox Checkbox)
        {
            original = Formula.original;
            shouldbe = Formula.shouldbe;
            worksheet = Formula.worksheet;
            condition = Formula.condition;
            location = Formula.location;

            checkbox = Checkbox;
        }

    }

    public partial class VerifyTests : Form
    {
        public List<TestCheck> TestsChecked = new List<TestCheck>();

        Expector Expector;

        public VerifyTests(Expector t)
        {
            Expector = t;
            InitializeComponent();
        }

        internal void PrintTest(string Text, testFormula Formula)
        {
            int NTestsPrinted = TestsChecked.Count;

            int height = 41 + 23 * NTestsPrinted;

            Label l = new Label();
            l.Location = new Point(12, height);
            l.Text = Formula.worksheet + "!"+Formula.location + ":"+ Text;
            l.AutoSize = true;
            this.Controls.Add(l);

            CheckBox C = new CheckBox();
            C.Checked = true;
            C.Location = new Point(260, height);
            this.Controls.Add(C);

            TestCheck t = new TestCheck(Formula, C);

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
	        }

            Expector.SaveTests(formulas);

            this.Close();
        }
    }
}
