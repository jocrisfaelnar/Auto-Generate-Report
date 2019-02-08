namespace WindowsFormsApp1
{
    partial class Form1
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
            this.btnMorning = new System.Windows.Forms.Button();
            this.btnEvening = new System.Windows.Forms.Button();
            this.cbTeamSelector = new System.Windows.Forms.ComboBox();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // btnMorning
            // 
            this.btnMorning.Location = new System.Drawing.Point(12, 87);
            this.btnMorning.Name = "btnMorning";
            this.btnMorning.Size = new System.Drawing.Size(75, 23);
            this.btnMorning.TabIndex = 0;
            this.btnMorning.Text = "Morning Report";
            this.btnMorning.Click += new System.EventHandler(this.btnMorning_Click);
            // 
            // btnEvening
            // 
            this.btnEvening.Location = new System.Drawing.Point(105, 87);
            this.btnEvening.Name = "btnEvening";
            this.btnEvening.Size = new System.Drawing.Size(75, 23);
            this.btnEvening.TabIndex = 1;
            this.btnEvening.Text = "Evening Report";
            this.btnEvening.UseVisualStyleBackColor = true;
            this.btnEvening.Click += new System.EventHandler(this.btnEvening_Click);
            // 
            // cbTeamSelector
            // 
            this.cbTeamSelector.FormattingEnabled = true;
            this.cbTeamSelector.Items.AddRange(new object[] {
            "<Empty>",
            "KX Report",
            "Collab Report"});
            this.cbTeamSelector.Location = new System.Drawing.Point(12, 51);
            this.cbTeamSelector.Name = "cbTeamSelector";
            this.cbTeamSelector.Size = new System.Drawing.Size(168, 21);
            this.cbTeamSelector.TabIndex = 2;
            this.cbTeamSelector.Text = "Select Team";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(134, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(46, 23);
            this.button1.TabIndex = 3;
            this.button1.Text = "Help?";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(12, 12);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(107, 22);
            this.button2.TabIndex = 4;
            this.button2.Text = "Edit APM Members";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(191, 124);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.cbTeamSelector);
            this.Controls.Add(this.btnEvening);
            this.Controls.Add(this.btnMorning);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SnowReportGenerator";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnMorning;
        private System.Windows.Forms.Button btnEvening;
        private System.Windows.Forms.ComboBox cbTeamSelector;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
    }
}

