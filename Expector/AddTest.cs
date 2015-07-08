using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Infotron.PerfectXL.DataModel;

namespace Expector
{
    public partial class AddTest : Form
    {
        Expector instanceofExpector;
        string _formula;
        string _worksheet;
        string _location;

        public AddTest(Expector e, string worksheet, string formula, string location)
        {
            instanceofExpector = e;
            _worksheet = worksheet;
            _formula = formula;
            _location = location;
            InitializeComponent();
            cellToAddTestsForLabel.Text = String.Format("You could a a test for the cell on {0}: {1}", _worksheet + "!" + _location, _formula);


        }

        private void AddTest_Load(object sender, EventArgs e)
        {
            comboBox1.DisplayMember = "Text";
            comboBox1.ValueMember = "Value";

            var items = new[] { 
                new { Text = "should be a number", Value = String.Format("ISNUMBER({0})",_location) }, 
                new { Text = "should not be a number", Value = String.Format("NOT(ISNUMBER({0}))",_location) }, 
                new { Text = "should be text", Value = String.Format("ISTEXT({0})",_location) },
                new { Text = "should not be text", Value = String.Format("NOT(ISTEXT({0}))",_location) },
                new { Text = "should not be blank", Value = String.Format("NOT(ISBLANK({0}))",_location) }
            };

            comboBox1.DataSource = items;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (typeCheck.Checked)
            {
                testFormula f = newTestForCurrentCell();
                f.condition = (string)comboBox1.SelectedValue;
                instanceofExpector.TestFormulas.Add(f);
            }

            if (lowerCheck.Checked)
            {
                testFormula f = newTestForCurrentCell();
                f.condition = String.Format("{0} > {1}", _location, lowerText.Text);
                instanceofExpector.TestFormulas.Add(f);
            }

            if (upperCheck.Checked)
            {
                testFormula f = newTestForCurrentCell();
                f.condition = String.Format("{0} < {1}", _location, upperText.Text);
                instanceofExpector.TestFormulas.Add(f);
            }

            instanceofExpector.SaveTests();

            MessageBox.Show("Test(s) added!");
            this.Close();
        }

        private testFormula newTestForCurrentCell()
        {
            testFormula f = new testFormula()
            {
                worksheet = _worksheet,
                location = _location,
            };
            return f;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            this.Height = 362;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
