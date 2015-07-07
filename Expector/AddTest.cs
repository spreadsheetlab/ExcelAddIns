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
        Cell cellToAddTestsFor;
        string location;

        public AddTest(Expector e, Cell c)
        {
            instanceofExpector = e;
            cellToAddTestsFor = c;
            location = cellToAddTestsFor.GetLocationIncludingSheetnameString(); //just a shorter way to get the name
            InitializeComponent();
            cellToAddTestsForLabel.Text = cellToAddTestsFor.GetLocationIncludingSheetnameString() + ":" + cellToAddTestsFor.Formula;
        }

        private void AddTest_Load(object sender, EventArgs e)
        {
            comboBox1.DisplayMember = "Text";
            comboBox1.ValueMember = "Value";

            var items = new[] { 
                new { Text = "should be a number", Value = String.Format("ISNUMBER({0})",location) }, 
                new { Text = "should not be a number", Value = String.Format("NOT(ISNUMBER({0}))",location) }, 
                new { Text = "should be text", Value = String.Format("ISTEXT({0})",location) },
                new { Text = "should not be text", Value = String.Format("NOT(ISTEXT({0}))",location) },
                new { Text = "should not be blank", Value = String.Format("NOT(ISBLANK({0}))",location) }
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
                f.condition = String.Format("{0} > {1}", location, lowerText.Text);
                instanceofExpector.TestFormulas.Add(f);
            }

            if (upperCheck.Checked)
            {
                testFormula f = newTestForCurrentCell();
                f.condition = String.Format("{0} < {1}", location, upperText.Text);
                instanceofExpector.TestFormulas.Add(f);
            }

            instanceofExpector.SaveTests();
            this.Close();
        }

        private testFormula newTestForCurrentCell()
        {
            testFormula f = new testFormula()
            {
                worksheet = cellToAddTestsFor.Worksheet.Name,
                location = cellToAddTestsFor.Location.ToString(),
            };
            return f;
        }
    }
}
