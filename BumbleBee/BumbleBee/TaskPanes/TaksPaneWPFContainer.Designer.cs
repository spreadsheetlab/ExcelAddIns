﻿namespace BumbleBee.TaskPanes
{
    partial class TaksPaneWPFContainer<T>
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.wpfElementHost = new System.Windows.Forms.Integration.ElementHost();
            this.SuspendLayout();
            // 
            // wpfElementHost
            // 
            this.wpfElementHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.wpfElementHost.Location = new System.Drawing.Point(0, 0);
            this.wpfElementHost.Name = "wpfElementHost";
            this.wpfElementHost.Size = new System.Drawing.Size(0, 0);
            this.wpfElementHost.TabIndex = 0;
            this.wpfElementHost.Text = "wpfElementHost";
            this.wpfElementHost.Child = null;
            // 
            // TaksPaneWPFContainer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(this.wpfElementHost);
            this.Name = "TaksPaneWPFContainer";
            this.Size = new System.Drawing.Size(0, 0);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Integration.ElementHost wpfElementHost;
    }
}
