namespace Netbattle.Forms
{
    partial class AdvancedConnect
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
            this.Command3 = new System.Windows.Forms.Button();
            this.Mode_4 = new System.Windows.Forms.RadioButton();
            this.Mode_3 = new System.Windows.Forms.RadioButton();
            this.Mode_2 = new System.Windows.Forms.RadioButton();
            this.Address = new System.Windows.Forms.ComboBox();
            this.Command2 = new System.Windows.Forms.Button();
            this.Command1 = new System.Windows.Forms.Button();
            this.Mode_1 = new System.Windows.Forms.RadioButton();
            this.Mode_0 = new System.Windows.Forms.RadioButton();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Command3
            // 
            this.Command3.Location = new System.Drawing.Point(8, 128);
            this.Command3.Name = "Command3";
            this.Command3.Size = new System.Drawing.Size(113, 25);
            this.Command3.TabIndex = 11;
            this.Command3.Text = "Clear &History";
            this.Command3.UseVisualStyleBackColor = true;
            // 
            // Mode_4
            // 
            this.Mode_4.Location = new System.Drawing.Point(144, 56);
            this.Mode_4.Name = "Mode_4";
            this.Mode_4.Size = new System.Drawing.Size(121, 17);
            this.Mode_4.TabIndex = 10;
            this.Mode_4.Enabled = false;
            this.Mode_4.Text = "S&ingle Player (vs. AI)";
            this.Mode_4.UseVisualStyleBackColor = true;
            // 
            // Mode_3
            // 
            this.Mode_3.Location = new System.Drawing.Point(144, 24);
            this.Mode_3.Name = "Mode_3";
            this.Mode_3.Size = new System.Drawing.Size(113, 17);
            this.Mode_3.TabIndex = 8;
            this.Mode_3.Enabled = false;
            this.Mode_3.Text = "&Host";
            this.Mode_3.UseVisualStyleBackColor = true;
            // 
            // Mode_2
            // 
            this.Mode_2.Location = new System.Drawing.Point(144, 40);
            this.Mode_2.Name = "Mode_2";
            this.Mode_2.Size = new System.Drawing.Size(113, 17);
            this.Mode_2.TabIndex = 7;
            this.Mode_2.Enabled = false;
            this.Mode_2.Text = "&Client";
            this.Mode_2.UseVisualStyleBackColor = true;
            // 
            // Address
            // 
            this.Address.Location = new System.Drawing.Point(8, 96);
            this.Address.Name = "Address";
            this.Address.Size = new System.Drawing.Size(257, 21);
            this.Address.TabIndex = 6;
            this.Address.FormattingEnabled = true;
            // 
            // Command2
            // 
            this.Command2.Location = new System.Drawing.Point(128, 128);
            this.Command2.Name = "Command2";
            this.Command2.Size = new System.Drawing.Size(65, 25);
            this.Command2.TabIndex = 4;
            this.Command2.Text = "&Cancel";
            this.Command2.UseVisualStyleBackColor = true;
            // 
            // Command1
            // 
            this.Command1.Location = new System.Drawing.Point(200, 128);
            this.Command1.Name = "Command1";
            this.Command1.Size = new System.Drawing.Size(65, 25);
            this.Command1.TabIndex = 3;
            this.Command1.Text = "&Ok";
            this.Command1.UseVisualStyleBackColor = true;
            // 
            // Mode_1
            // 
            this.Mode_1.Location = new System.Drawing.Point(8, 40);
            this.Mode_1.Name = "Mode_1";
            this.Mode_1.Size = new System.Drawing.Size(105, 17);
            this.Mode_1.TabIndex = 1;
            this.Mode_1.Checked = true;
            this.Mode_1.Text = "S&pecific Server";
            this.Mode_1.UseVisualStyleBackColor = true;
            // 
            // Mode_0
            // 
            this.Mode_0.Location = new System.Drawing.Point(8, 24);
            this.Mode_0.Name = "Mode_0";
            this.Mode_0.Size = new System.Drawing.Size(105, 17);
            this.Mode_0.TabIndex = 0;
            this.Mode_0.Text = "&Server Browser";
            this.Mode_0.UseVisualStyleBackColor = true;
            // 
            // Label3
            // 
            this.Label3.Location = new System.Drawing.Point(144, 8);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(89, 17);
            this.Label3.TabIndex = 9;
            this.Label3.BackColor = System.Drawing.Color.Transparent;
            this.Label3.Text = "Direct Connect:";
            // 
            // Label2
            // 
            this.Label2.Location = new System.Drawing.Point(8, 8);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(113, 17);
            this.Label2.TabIndex = 5;
            this.Label2.BackColor = System.Drawing.Color.Transparent;
            this.Label2.Text = "Connect to:";
            // 
            // Label1
            // 
            this.Label1.Location = new System.Drawing.Point(8, 80);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(137, 17);
            this.Label1.TabIndex = 2;
            this.Label1.BackColor = System.Drawing.Color.Transparent;
            this.Label1.Text = "Server (Name or IP Address)";
            // 
            // AdvancedConnect
            // 
            this.AcceptButton = this.Command1;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.Command2;
            this.ClientSize = new System.Drawing.Size(274, 161);
            this.ControlBox = false;
            this.Controls.Add(this.Command3);
            this.Controls.Add(this.Mode_4);
            this.Controls.Add(this.Mode_3);
            this.Controls.Add(this.Mode_2);
            this.Controls.Add(this.Address);
            this.Controls.Add(this.Command2);
            this.Controls.Add(this.Command1);
            this.Controls.Add(this.Mode_1);
            this.Controls.Add(this.Mode_0);
            this.Controls.Add(this.Label3);
            this.Controls.Add(this.Label2);
            this.Controls.Add(this.Label1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AdvancedConnect";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Advanced Connect";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Button Command3;
        private System.Windows.Forms.RadioButton Mode_4;
        private System.Windows.Forms.RadioButton Mode_3;
        private System.Windows.Forms.RadioButton Mode_2;
        private System.Windows.Forms.ComboBox Address;
        private System.Windows.Forms.Button Command2;
        private System.Windows.Forms.Button Command1;
        private System.Windows.Forms.RadioButton Mode_1;
        private System.Windows.Forms.RadioButton Mode_0;
        private System.Windows.Forms.Label Label3;
        private System.Windows.Forms.Label Label2;
        private System.Windows.Forms.Label Label1;
    }
}
