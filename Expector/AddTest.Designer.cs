namespace Expector
{
    partial class AddTest
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.typeCheck = new System.Windows.Forms.CheckBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.lowerCheck = new System.Windows.Forms.CheckBox();
            this.lowerText = new System.Windows.Forms.TextBox();
            this.upperText = new System.Windows.Forms.TextBox();
            this.upperCheck = new System.Windows.Forms.CheckBox();
            this.saveAddedTests = new System.Windows.Forms.Button();
            this.cellToAddTestsForLabel = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // typeCheck
            // 
            this.typeCheck.AutoSize = true;
            this.typeCheck.Location = new System.Drawing.Point(12, 119);
            this.typeCheck.Name = "typeCheck";
            this.typeCheck.Size = new System.Drawing.Size(68, 17);
            this.typeCheck.TabIndex = 0;
            this.typeCheck.Text = "This cell ";
            this.typeCheck.UseVisualStyleBackColor = true;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(151, 117);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 21);
            this.comboBox1.TabIndex = 1;
            // 
            // lowerCheck
            // 
            this.lowerCheck.AutoSize = true;
            this.lowerCheck.Location = new System.Drawing.Point(12, 155);
            this.lowerCheck.Name = "lowerCheck";
            this.lowerCheck.Size = new System.Drawing.Size(135, 17);
            this.lowerCheck.TabIndex = 2;
            this.lowerCheck.Text = "Value should be above";
            this.lowerCheck.UseVisualStyleBackColor = true;
            // 
            // lowerText
            // 
            this.lowerText.Location = new System.Drawing.Point(151, 155);
            this.lowerText.Name = "lowerText";
            this.lowerText.Size = new System.Drawing.Size(51, 20);
            this.lowerText.TabIndex = 3;
            // 
            // upperText
            // 
            this.upperText.Location = new System.Drawing.Point(151, 192);
            this.upperText.Name = "upperText";
            this.upperText.Size = new System.Drawing.Size(51, 20);
            this.upperText.TabIndex = 5;
            // 
            // upperCheck
            // 
            this.upperCheck.AutoSize = true;
            this.upperCheck.Location = new System.Drawing.Point(12, 192);
            this.upperCheck.Name = "upperCheck";
            this.upperCheck.Size = new System.Drawing.Size(133, 17);
            this.upperCheck.TabIndex = 4;
            this.upperCheck.Text = "Value should be below";
            this.upperCheck.UseVisualStyleBackColor = true;
            // 
            // saveAddedTests
            // 
            this.saveAddedTests.Location = new System.Drawing.Point(12, 278);
            this.saveAddedTests.Name = "saveAddedTests";
            this.saveAddedTests.Size = new System.Drawing.Size(75, 23);
            this.saveAddedTests.TabIndex = 6;
            this.saveAddedTests.Text = "Save";
            this.saveAddedTests.UseVisualStyleBackColor = true;
            this.saveAddedTests.Click += new System.EventHandler(this.button1_Click);
            // 
            // cellToAddTestsForLabel
            // 
            this.cellToAddTestsForLabel.AutoSize = true;
            this.cellToAddTestsForLabel.Location = new System.Drawing.Point(13, 27);
            this.cellToAddTestsForLabel.Name = "cellToAddTestsForLabel";
            this.cellToAddTestsForLabel.Size = new System.Drawing.Size(0, 13);
            this.cellToAddTestsForLabel.TabIndex = 7;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(151, 231);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(199, 20);
            this.textBox3.TabIndex = 9;
            this.textBox3.Visible = false;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(12, 231);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(125, 17);
            this.checkBox1.TabIndex = 8;
            this.checkBox1.Text = "Enter custom formula";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.Visible = false;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 74);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(105, 23);
            this.button1.TabIndex = 10;
            this.button1.Text = "Yes, make tests";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(127, 74);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 11;
            this.button2.Text = "No, thanks";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // AddTest
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(406, 106);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.cellToAddTestsForLabel);
            this.Controls.Add(this.saveAddedTests);
            this.Controls.Add(this.upperText);
            this.Controls.Add(this.upperCheck);
            this.Controls.Add(this.lowerText);
            this.Controls.Add(this.lowerCheck);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.typeCheck);
            this.Name = "AddTest";
            this.Text = "Add new tests";
            this.Load += new System.EventHandler(this.AddTest_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox typeCheck;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.CheckBox lowerCheck;
        private System.Windows.Forms.TextBox lowerText;
        private System.Windows.Forms.TextBox upperText;
        private System.Windows.Forms.CheckBox upperCheck;
        private System.Windows.Forms.Button saveAddedTests;
        private System.Windows.Forms.Label cellToAddTestsForLabel;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}