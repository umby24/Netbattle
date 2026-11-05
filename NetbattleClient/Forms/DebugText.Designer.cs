namespace Netbattle.Forms
{
    partial class DebugText
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
            this.RichTextBox1 = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // RichTextBox1
            // 
            this.RichTextBox1.Location = new System.Drawing.Point(16, 16);
            this.RichTextBox1.Name = "RichTextBox1";
            this.RichTextBox1.Size = new System.Drawing.Size(273, 161);
            this.RichTextBox1.TabIndex = 0;
            this.RichTextBox1.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None;
            this.RichTextBox1.Text = "";
            // 
            // DebugText
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(312, 206);
            this.Controls.Add(this.RichTextBox1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "DebugText";
            this.Text = "Form2";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.RichTextBox RichTextBox1;
    }
}
