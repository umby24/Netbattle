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
        private void InitializeComponent() {
            MoveDown = new System.Windows.Forms.Button();
            MoveUp = new System.Windows.Forms.Button();
            List1 = new System.Windows.Forms.ListBox();
            CancelButton_ = new System.Windows.Forms.Button();
            OKButton = new System.Windows.Forms.Button();
            Label2 = new System.Windows.Forms.Label();
            Label1 = new System.Windows.Forms.Label();
            SuspendLayout();
            // 
            // MoveDown
            // 
            MoveDown.Font = new System.Drawing.Font("Wingdings", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)2));
            MoveDown.Location = new System.Drawing.Point(5, 47);
            MoveDown.Name = "MoveDown";
            MoveDown.Size = new System.Drawing.Size(21, 22);
            MoveDown.TabIndex = 4;
            MoveDown.Text = "ê";
            MoveDown.UseVisualStyleBackColor = true;
            MoveDown.Click += MoveDown_Click;
            // 
            // MoveUp
            // 
            MoveUp.Font = new System.Drawing.Font("Wingdings", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)2));
            MoveUp.Location = new System.Drawing.Point(5, 16);
            MoveUp.Name = "MoveUp";
            MoveUp.Size = new System.Drawing.Size(21, 22);
            MoveUp.TabIndex = 3;
            MoveUp.Text = "é";
            MoveUp.UseVisualStyleBackColor = true;
            MoveUp.Click += MoveUp_Click;
            // 
            // List1
            // 
            List1.FormattingEnabled = true;
            List1.Location = new System.Drawing.Point(30, 16);
            List1.Name = "List1";
            List1.Size = new System.Drawing.Size(128, 56);
            List1.TabIndex = 2;
            // 
            // CancelButton_
            // 
            CancelButton_.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            CancelButton_.Location = new System.Drawing.Point(105, 73);
            CancelButton_.Name = "CancelButton_";
            CancelButton_.Size = new System.Drawing.Size(51, 16);
            CancelButton_.TabIndex = 1;
            CancelButton_.Text = "Cancel";
            CancelButton_.UseVisualStyleBackColor = true;
            CancelButton_.Click += CancelButton__Click;
            // 
            // OKButton
            // 
            OKButton.Location = new System.Drawing.Point(51, 73);
            OKButton.Name = "OKButton";
            OKButton.Size = new System.Drawing.Size(51, 16);
            OKButton.TabIndex = 0;
            OKButton.Text = "OK";
            OKButton.UseVisualStyleBackColor = true;
            OKButton.Click += OKButton_Click;
            // 
            // Label2
            // 
            Label2.BackColor = System.Drawing.Color.Transparent;
            Label2.Location = new System.Drawing.Point(30, 5);
            Label2.Name = "Label2";
            Label2.Size = new System.Drawing.Size(126, 11);
            Label2.TabIndex = 6;
            Label2.Text = "Current Order";
            // 
            // Label1
            // 
            Label1.BackColor = System.Drawing.Color.Transparent;
            Label1.Location = new System.Drawing.Point(5, 5);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(21, 11);
            Label1.TabIndex = 5;
            Label1.Text = "Move";
            // 
            // RearrangeTeam
            // 
            AcceptButton = OKButton;
            AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            CancelButton = CancelButton_;
            ClientSize = new System.Drawing.Size(162, 94);
            Controls.Add(MoveDown);
            Controls.Add(MoveUp);
            Controls.Add(List1);
            Controls.Add(CancelButton_);
            Controls.Add(OKButton);
            Controls.Add(Label2);
            Controls.Add(Label1);
            Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)0));
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            Location = new System.Drawing.Point(184, 242);
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            Text = "Rearrange Team";
            ResumeLayout(false);
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
