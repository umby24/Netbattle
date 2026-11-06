using System.ComponentModel;

namespace Netbattle.Forms
{
    partial class TempBanSet
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
            ComponentResourceManager resources = new ComponentResourceManager(typeof(TempBanSet));
            btnCancel = new Button();
            btnOk = new Button();
            label1 = new Label();
            duration = new NumericUpDown();
            ((ISupportInitialize)duration).BeginInit();
            SuspendLayout();
            // 
            // btnCancel
            // 
            btnCancel.Location = new Point(10, 68);
            btnCancel.Margin = new Padding(4, 3, 4, 3);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(58, 28);
            btnCancel.TabIndex = 0;
            btnCancel.Text = "&Cancel";
            btnCancel.UseVisualStyleBackColor = true;
            btnCancel.Click += btnCancel_Click;
            // 
            // btnOk
            // 
            btnOk.Location = new Point(76, 69);
            btnOk.Margin = new Padding(4, 3, 4, 3);
            btnOk.Name = "btnOk";
            btnOk.Size = new Size(58, 28);
            btnOk.TabIndex = 1;
            btnOk.Text = "&OK";
            btnOk.UseVisualStyleBackColor = true;
            btnOk.Click += btnOk_Click;
            // 
            // label1
            // 
            label1.Location = new Point(10, 10);
            label1.Margin = new Padding(4, 0, 4, 0);
            label1.Name = "label1";
            label1.Size = new Size(105, 20);
            label1.TabIndex = 2;
            label1.Text = "Enter duration:";
            // 
            // duration
            // 
            duration.Location = new Point(14, 33);
            duration.Margin = new Padding(4, 3, 4, 3);
            duration.Maximum = new decimal(new int[] { 9999, 0, 0, 0 });
            duration.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            duration.Name = "duration";
            duration.Size = new Size(71, 23);
            duration.TabIndex = 3;
            duration.Value = new decimal(new int[] { 15, 0, 0, 0 });
            // 
            // TempBanSet
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(138, 97);
            Controls.Add(duration);
            Controls.Add(label1);
            Controls.Add(btnOk);
            Controls.Add(btnCancel);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Margin = new Padding(4, 3, 4, 3);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "TempBanSet";
            Text = "Temp Ban";
            ((ISupportInitialize)duration).EndInit();
            ResumeLayout(false);
        }

        private System.Windows.Forms.NumericUpDown duration;

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Label label1;

        #endregion
    }
}