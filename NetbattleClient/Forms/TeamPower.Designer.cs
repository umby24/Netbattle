namespace Netbattle.Forms
{
    partial class TeamPower
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
            this.Command1 = new System.Windows.Forms.Button();
            this.Text1 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // Command1
            // 
            this.Command1.Location = new System.Drawing.Point(88, 8);
            this.Command1.Name = "Command1";
            this.Command1.Size = new System.Drawing.Size(73, 17);
            this.Command1.TabIndex = 1;
            this.Command1.Text = "Change";
            this.Command1.UseVisualStyleBackColor = true;
            // 
            // Text1
            // 
            this.Text1.Location = new System.Drawing.Point(8, 8);
            this.Text1.Name = "Text1";
            this.Text1.Size = new System.Drawing.Size(65, 19);
            this.Text1.TabIndex = 0;
            this.Text1.Text = "Text1";
            // 
            // TeamPower
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(172, 35);
            this.Controls.Add(this.Command1);
            this.Controls.Add(this.Text1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "TeamPower";
            this.ShowInTaskbar = false;
            this.Text = "Change Team Power";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Button Command1;
        private System.Windows.Forms.TextBox Text1;
    }
}
