namespace Expector
{
    partial class VerifyTests
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
            this.Tests = new System.Windows.Forms.Label();
            this.IsTestCheckBox = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Tests
            // 
            this.Tests.Location = new System.Drawing.Point(0, 0);
            this.Tests.Name = "Tests";
            this.Tests.Size = new System.Drawing.Size(100, 23);
            this.Tests.TabIndex = 7;
            // 
            // IsTestCheckBox
            // 
            this.IsTestCheckBox.Location = new System.Drawing.Point(0, 0);
            this.IsTestCheckBox.Name = "IsTestCheckBox";
            this.IsTestCheckBox.Size = new System.Drawing.Size(104, 24);
            this.IsTestCheckBox.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 23);
            this.label1.TabIndex = 5;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "Save tests";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // VerifyTests
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(490, 831);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.IsTestCheckBox);
            this.Controls.Add(this.Tests);
            this.Name = "VerifyTests";
            this.Text = "Tests Detected";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label Tests;
        private System.Windows.Forms.CheckBox IsTestCheckBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
    }
}