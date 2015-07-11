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
        string _totalLocation;

        public AddTest(Expector e, string worksheet, string formula, string location)
        {
            instanceofExpector = e;
            _worksheet = worksheet;
            _formula = formula;
            _location = location;
            _totalLocation = _worksheet + "!" + _location;
            InitializeComponent();
            cellToAddTestsForLabel.Text = String.Format("You could a a test for the cell on {0}: {1}", _totalLocation, _formula);


        }

        private void AddTest_Load(object sender, EventArgs e)
        {
            comboBox1.DisplayMember = "Text";
            comboBox1.ValueMember = "Value";

            var items = new[] { 
                new { Text = "should be a number", Value = String.Format("ISNUMBER({0})",_totalLocation) }, 
                new { Text = "should not be a number", Value = String.Format("NOT(ISNUMBER({0}))",_totalLocation) }, 
                new { Text = "should be text", Value = String.Format("ISTEXT({0})",_totalLocation) },
                new { Text = "should not be text", Value = String.Format("NOT(ISTEXT({0}))",_totalLocation) },
                new { Text = "should not be blank", Value = String.Format("NOT(ISBLANK({0}))",_totalLocation) }
            };

            comboBox1.DataSource = items;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double coverageBefore = instanceofExpector.getCurrentCoverage();
            bool validTestsFound = false;

            if (typeCheck.Checked)
            {
                testFormula f = newTestForCurrentCell();
                f.condition = (string)comboBox1.SelectedValue;
                instanceofExpector.testFormulas.Add(f);
                validTestsFound = true;
            }

            if (lowerCheck.Checked)
            {
                if (lowerText.Text != "" )
                {
                    testFormula f = newTestForCurrentCell();
                    f.condition = String.Format("{0} > {1}", _totalLocation, lowerText.Text);
                    instanceofExpector.testFormulas.Add(f);
                    validTestsFound = true;
                }
                else
                {
                    MessageBox.Show("No bound filled out, please input one.");
                }
            }

            if (upperCheck.Checked)
            {
                if (upperText.Text != "")
                {
                    testFormula f = newTestForCurrentCell();
                    f.condition = String.Format("{0} < {1}", _totalLocation, upperText.Text);
                    instanceofExpector.testFormulas.Add(f);
                    validTestsFound = true;
                }
                else
                {
                    MessageBox.Show("No bound filled out, please input one.");
                }
            }

            if (typeCheck.Checked || lowerCheck.Checked || upperCheck.Checked)
            {
                if (validTestsFound)
                {
                    instanceofExpector.SaveTests();

                    //update the covered and non covered cells cells
                    instanceofExpector.initCellsLists();

                    double coverageAfter = instanceofExpector.getCurrentCoverage();

                    MessageBox.Show(String.Format("Wonderful, you have increased coverage from {0}% to {1}%", Math.Round(coverageBefore), Math.Round(coverageAfter)));
                    this.Close();
                }
            }
            else
            {
                MessageBox.Show("No tests selected");
            }




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
