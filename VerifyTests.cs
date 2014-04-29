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
    public class TestCheck
    {
        public string formula;
        public CheckBox checkbox;

        public TestCheck(string Formula, CheckBox Checkbox)
        {
            formula = Formula;
            checkbox = Checkbox;
        }

    }

    public partial class VerifyTests : Form
    {
        public List<TestCheck> TestsChecked = new List<TestCheck>();

        ThisAddIn Expector;

        public VerifyTests(ThisAddIn t)
        {
            Expector = t;
            InitializeComponent();
        }

        internal void PrintTest(string Text, string Formula)
        {
            int NTestsPrinted = TestsChecked.Count;

            int height = 41 + 23 * NTestsPrinted;

            Label l = new Label();
            l.Location = new Point(12, height);
            l.Text = Text;
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
            List<string> formulas = new List<string>();

            foreach (var c in TestsChecked)
	        {
                int i = 0;
		        if (c.checkbox.Checked)
	            {
                    formulas.Add(c.formula);
                    i++;
	            }
	        }

            MessageBox.Show("Tests saved");
            Expector.SaveTests(formulas);

            this.Close();
        }
    }
}
