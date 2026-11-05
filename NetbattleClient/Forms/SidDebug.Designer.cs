namespace Netbattle.Forms
{
    partial class SidDebug
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
            this.Command2 = new System.Windows.Forms.Button();
            this.Command1 = new System.Windows.Forms.Button();
            this.Text1 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // Command2
            // 
            this.Command2.Location = new System.Drawing.Point(248, 0);
            this.Command2.Name = "Command2";
            this.Command2.Size = new System.Drawing.Size(49, 25);
            this.Command2.TabIndex = 2;
            this.Command2.Text = "Normal";
            this.Command2.UseVisualStyleBackColor = true;
            // 
            // Command1
            // 
            this.Command1.Location = new System.Drawing.Point(184, 0);
            this.Command1.Name = "Command1";
            this.Command1.Size = new System.Drawing.Size(57, 25);
            this.Command1.TabIndex = 1;
            this.Command1.Text = "Set";
            this.Command1.UseVisualStyleBackColor = true;
            // 
            // Text1
            // 
            this.Text1.Location = new System.Drawing.Point(8, 0);
            this.Text1.Name = "Text1";
            this.Text1.Size = new System.Drawing.Size(169, 25);
            this.Text1.TabIndex = 0;
            this.Text1.MaxLength = 13;
            this.Text1.Text = "Text1";
            // 
            // SidDebug
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(312, 33);
            this.Controls.Add(this.Command2);
            this.Controls.Add(this.Command1);
            this.Controls.Add(this.Text1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SidDebug";
            this.Text = "SID";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Button Command2;
        private System.Windows.Forms.Button Command1;
        private System.Windows.Forms.TextBox Text1;
    }
}
