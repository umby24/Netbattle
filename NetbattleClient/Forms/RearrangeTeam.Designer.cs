namespace Netbattle.Forms
{
    partial class RearrangeTeam
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
            this.MoveDown = new System.Windows.Forms.Button();
            this.MoveUp = new System.Windows.Forms.Button();
            this.List1 = new System.Windows.Forms.ListBox();
            this.CancelButton_ = new System.Windows.Forms.Button();
            this.OKButton = new System.Windows.Forms.Button();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // MoveDown
            // 
            this.MoveDown.Font = new System.Drawing.Font("Wingdings", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.MoveDown.Location = new System.Drawing.Point(8, 72);
            this.MoveDown.Name = "MoveDown";
            this.MoveDown.Size = new System.Drawing.Size(33, 33);
            this.MoveDown.TabIndex = 4;
            this.MoveDown.Text = "ê";
            this.MoveDown.UseVisualStyleBackColor = true;
            // 
            // MoveUp
            // 
            this.MoveUp.Font = new System.Drawing.Font("Wingdings", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.MoveUp.Location = new System.Drawing.Point(8, 24);
            this.MoveUp.Name = "MoveUp";
            this.MoveUp.Size = new System.Drawing.Size(33, 33);
            this.MoveUp.TabIndex = 3;
            this.MoveUp.Text = "é";
            this.MoveUp.UseVisualStyleBackColor = true;
            // 
            // List1
            // 
            this.List1.Location = new System.Drawing.Point(48, 24);
            this.List1.Name = "List1";
            this.List1.Size = new System.Drawing.Size(201, 82);
            this.List1.TabIndex = 2;
            this.List1.FormattingEnabled = true;
            // 
            // CancelButton_
            // 
            this.CancelButton_.Location = new System.Drawing.Point(168, 112);
            this.CancelButton_.Name = "CancelButton_";
            this.CancelButton_.Size = new System.Drawing.Size(81, 25);
            this.CancelButton_.TabIndex = 1;
            this.CancelButton_.Text = "Cancel";
            this.CancelButton_.UseVisualStyleBackColor = true;
            // 
            // OKButton
            // 
            this.OKButton.Location = new System.Drawing.Point(80, 112);
            this.OKButton.Name = "OKButton";
            this.OKButton.Size = new System.Drawing.Size(81, 25);
            this.OKButton.TabIndex = 0;
            this.OKButton.Text = "OK";
            this.OKButton.UseVisualStyleBackColor = true;
            // 
            // Label2
            // 
            this.Label2.Location = new System.Drawing.Point(48, 8);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(201, 17);
            this.Label2.TabIndex = 6;
            this.Label2.BackColor = System.Drawing.Color.Transparent;
            this.Label2.Text = "Current Order";
            // 
            // Label1
            // 
            this.Label1.Location = new System.Drawing.Point(8, 8);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(33, 17);
            this.Label1.TabIndex = 5;
            this.Label1.BackColor = System.Drawing.Color.Transparent;
            this.Label1.Text = "Move";
            // 
            // RearrangeTeam
            // 
            this.AcceptButton = this.OKButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CancelButton_;
            this.ClientSize = new System.Drawing.Size(258, 145);
            this.Controls.Add(this.MoveDown);
            this.Controls.Add(this.MoveUp);
            this.Controls.Add(this.List1);
            this.Controls.Add(this.CancelButton_);
            this.Controls.Add(this.OKButton);
            this.Controls.Add(this.Label2);
            this.Controls.Add(this.Label1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Location = new System.Drawing.Point(184, 242);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "RearrangeTeam";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Rearrange Team";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Button MoveDown;
        private System.Windows.Forms.Button MoveUp;
        private System.Windows.Forms.ListBox List1;
        private System.Windows.Forms.Button CancelButton_;
        private System.Windows.Forms.Button OKButton;
        private System.Windows.Forms.Label Label2;
        private System.Windows.Forms.Label Label1;
    }
}
