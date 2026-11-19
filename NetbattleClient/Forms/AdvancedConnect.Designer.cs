namespace NetbattleClient.Forms
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
        private void InitializeComponent() {
            Command3 = new System.Windows.Forms.Button();
            Mode_4 = new System.Windows.Forms.RadioButton();
            Mode_3 = new System.Windows.Forms.RadioButton();
            Mode_2 = new System.Windows.Forms.RadioButton();
            Address = new System.Windows.Forms.ComboBox();
            Command2 = new System.Windows.Forms.Button();
            Command1 = new System.Windows.Forms.Button();
            Mode_1 = new System.Windows.Forms.RadioButton();
            Mode_0 = new System.Windows.Forms.RadioButton();
            Label3 = new System.Windows.Forms.Label();
            Label2 = new System.Windows.Forms.Label();
            Label1 = new System.Windows.Forms.Label();
            SuspendLayout();
            // 
            // Command3
            // 
            Command3.Location = new System.Drawing.Point(4, 142);
            Command3.Name = "Command3";
            Command3.Size = new System.Drawing.Size(90, 23);
            Command3.TabIndex = 11;
            Command3.Text = "Clear &History";
            Command3.UseVisualStyleBackColor = true;
            Command3.Click += Command3_Click;
            // 
            // Mode_4
            // 
            Mode_4.Enabled = false;
            Mode_4.Location = new System.Drawing.Point(108, 63);
            Mode_4.Name = "Mode_4";
            Mode_4.Size = new System.Drawing.Size(134, 16);
            Mode_4.TabIndex = 10;
            Mode_4.Text = "S&ingle Player (vs. AI)";
            Mode_4.UseVisualStyleBackColor = true;
            // 
            // Mode_3
            // 
            Mode_3.Enabled = false;
            Mode_3.Location = new System.Drawing.Point(108, 23);
            Mode_3.Name = "Mode_3";
            Mode_3.Size = new System.Drawing.Size(51, 21);
            Mode_3.TabIndex = 8;
            Mode_3.Text = "&Host";
            Mode_3.UseVisualStyleBackColor = true;
            // 
            // Mode_2
            // 
            Mode_2.Enabled = false;
            Mode_2.Location = new System.Drawing.Point(108, 42);
            Mode_2.Name = "Mode_2";
            Mode_2.Size = new System.Drawing.Size(80, 21);
            Mode_2.TabIndex = 7;
            Mode_2.Text = "&Client";
            Mode_2.UseVisualStyleBackColor = true;
            // 
            // Address
            // 
            Address.FormattingEnabled = true;
            Address.Location = new System.Drawing.Point(12, 115);
            Address.Name = "Address";
            Address.Size = new System.Drawing.Size(201, 21);
            Address.TabIndex = 6;
            // 
            // Command2
            // 
            Command2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            Command2.Location = new System.Drawing.Point(100, 142);
            Command2.Name = "Command2";
            Command2.Size = new System.Drawing.Size(67, 23);
            Command2.TabIndex = 4;
            Command2.Text = "&Cancel";
            Command2.UseVisualStyleBackColor = true;
            Command2.Click += Command2_Click;
            // 
            // Command1
            // 
            Command1.Location = new System.Drawing.Point(170, 142);
            Command1.Name = "Command1";
            Command1.Size = new System.Drawing.Size(61, 23);
            Command1.TabIndex = 3;
            Command1.Text = "&Ok";
            Command1.UseVisualStyleBackColor = true;
            Command1.Click += Command1_Click;
            // 
            // Mode_1
            // 
            Mode_1.Checked = true;
            Mode_1.Location = new System.Drawing.Point(4, 47);
            Mode_1.Name = "Mode_1";
            Mode_1.Size = new System.Drawing.Size(98, 16);
            Mode_1.TabIndex = 1;
            Mode_1.TabStop = true;
            Mode_1.Text = "S&pecific Server";
            Mode_1.UseVisualStyleBackColor = true;
            // 
            // Mode_0
            // 
            Mode_0.Location = new System.Drawing.Point(4, 20);
            Mode_0.Name = "Mode_0";
            Mode_0.Size = new System.Drawing.Size(98, 27);
            Mode_0.TabIndex = 0;
            Mode_0.Text = "&Server Browser";
            Mode_0.UseVisualStyleBackColor = true;
            // 
            // Label3
            // 
            Label3.BackColor = System.Drawing.Color.Transparent;
            Label3.Location = new System.Drawing.Point(108, 5);
            Label3.Name = "Label3";
            Label3.Size = new System.Drawing.Size(93, 19);
            Label3.TabIndex = 9;
            Label3.Text = "Direct Connect:";
            // 
            // Label2
            // 
            Label2.BackColor = System.Drawing.Color.Transparent;
            Label2.Location = new System.Drawing.Point(4, 4);
            Label2.Name = "Label2";
            Label2.Size = new System.Drawing.Size(80, 13);
            Label2.TabIndex = 5;
            Label2.Text = "Connect to:";
            // 
            // Label1
            // 
            Label1.BackColor = System.Drawing.Color.Transparent;
            Label1.Location = new System.Drawing.Point(12, 97);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(201, 15);
            Label1.TabIndex = 2;
            Label1.Text = "Server (Name or IP Address)";
            // 
            // AdvancedConnect
            // 
            AcceptButton = Command1;
            AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            CancelButton = Command2;
            ClientSize = new System.Drawing.Size(243, 173);
            ControlBox = false;
            Controls.Add(Command3);
            Controls.Add(Mode_4);
            Controls.Add(Mode_3);
            Controls.Add(Mode_2);
            Controls.Add(Address);
            Controls.Add(Command2);
            Controls.Add(Command1);
            Controls.Add(Mode_1);
            Controls.Add(Mode_0);
            Controls.Add(Label3);
            Controls.Add(Label2);
            Controls.Add(Label1);
            Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)0));
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Advanced Connect";
            Load += AdvancedConnect_Load;
            ResumeLayout(false);
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
