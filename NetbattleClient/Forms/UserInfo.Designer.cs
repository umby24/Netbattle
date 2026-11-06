namespace Netbattle.Forms
{
    partial class UserInfo
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
            SSTab1 = new TabControl();
            SSTab1_Page0 = new TabPage();
            Text1 = new TextBox();
            cmdLookup = new Button();
            Command_0 = new Button();
            txtInfo_0 = new TextBox();
            txtInfo_1 = new TextBox();
            txtInfo_2 = new TextBox();
            txtInfo_3 = new TextBox();
            Command_1 = new Button();
            Command_2 = new Button();
            txtInfo_4 = new TextBox();
            Label1 = new Label();
            Label2 = new Label();
            Label3 = new Label();
            Label4 = new Label();
            Label5 = new Label();
            SSTab1_Page1 = new TabPage();
            cmdRefresh = new Button();
            cmdUnban = new Button();
            BanList = new ListView();
            BanList_Col1 = new ColumnHeader();
            BanList_Col2 = new ColumnHeader();
            SSTab1_Page2 = new TabPage();
            cmdALook = new Button();
            Text2 = new TextBox();
            NickList = new ListView();
            NickList_Col1 = new ColumnHeader();
            StatusBar = new StatusStrip();
            StatusBar_Panel1 = new ToolStripStatusLabel();
            SSTab1.SuspendLayout();
            SSTab1_Page0.SuspendLayout();
            SSTab1_Page1.SuspendLayout();
            SSTab1_Page2.SuspendLayout();
            StatusBar.SuspendLayout();
            SuspendLayout();
            // 
            // SSTab1
            // 
            SSTab1.Controls.Add(SSTab1_Page0);
            SSTab1.Controls.Add(SSTab1_Page1);
            SSTab1.Controls.Add(SSTab1_Page2);
            SSTab1.Location = new Point(4, 4);
            SSTab1.Name = "SSTab1";
            SSTab1.SelectedIndex = 0;
            SSTab1.Size = new Size(241, 177);
            SSTab1.TabIndex = 1;
            // 
            // SSTab1_Page0
            // 
            SSTab1_Page0.Controls.Add(Text1);
            SSTab1_Page0.Controls.Add(cmdLookup);
            SSTab1_Page0.Controls.Add(Command_0);
            SSTab1_Page0.Controls.Add(txtInfo_0);
            SSTab1_Page0.Controls.Add(txtInfo_1);
            SSTab1_Page0.Controls.Add(txtInfo_2);
            SSTab1_Page0.Controls.Add(txtInfo_3);
            SSTab1_Page0.Controls.Add(Command_1);
            SSTab1_Page0.Controls.Add(Command_2);
            SSTab1_Page0.Controls.Add(txtInfo_4);
            SSTab1_Page0.Controls.Add(Label1);
            SSTab1_Page0.Controls.Add(Label2);
            SSTab1_Page0.Controls.Add(Label3);
            SSTab1_Page0.Controls.Add(Label4);
            SSTab1_Page0.Controls.Add(Label5);
            SSTab1_Page0.Location = new Point(4, 22);
            SSTab1_Page0.Name = "SSTab1_Page0";
            SSTab1_Page0.Size = new Size(233, 151);
            SSTab1_Page0.TabIndex = 0;
            SSTab1_Page0.Text = "User Lookup";
            SSTab1_Page0.UseVisualStyleBackColor = true;
            // 
            // Text1
            // 
            Text1.Location = new Point(4, 6);
            Text1.Name = "Text1";
            Text1.Size = new Size(169, 20);
            Text1.TabIndex = 11;
            // 
            // cmdLookup
            // 
            cmdLookup.Location = new Point(180, 6);
            cmdLookup.Name = "cmdLookup";
            cmdLookup.Size = new Size(49, 21);
            cmdLookup.TabIndex = 10;
            cmdLookup.Text = "Lookup";
            cmdLookup.UseVisualStyleBackColor = true;
            cmdLookup.Click += cmdLookup_Click;
            // 
            // Command_0
            // 
            Command_0.Enabled = false;
            Command_0.Location = new Point(4, 126);
            Command_0.Name = "Command_0";
            Command_0.Size = new Size(65, 21);
            Command_0.TabIndex = 9;
            Command_0.Text = "TempBan";
            Command_0.UseVisualStyleBackColor = true;
            Command_0.Click += Command_0_Click;
            // 
            // txtInfo_0
            // 
            txtInfo_0.BackColor = SystemColors.Control;
            txtInfo_0.BorderStyle = BorderStyle.None;
            txtInfo_0.Location = new Point(68, 38);
            txtInfo_0.Name = "txtInfo_0";
            txtInfo_0.ReadOnly = true;
            txtInfo_0.Size = new Size(161, 13);
            txtInfo_0.TabIndex = 8;
            txtInfo_0.Text = "[No User Selected]";
            // 
            // txtInfo_1
            // 
            txtInfo_1.BackColor = SystemColors.Control;
            txtInfo_1.BorderStyle = BorderStyle.None;
            txtInfo_1.Location = new Point(68, 54);
            txtInfo_1.Name = "txtInfo_1";
            txtInfo_1.ReadOnly = true;
            txtInfo_1.Size = new Size(161, 13);
            txtInfo_1.TabIndex = 7;
            // 
            // txtInfo_2
            // 
            txtInfo_2.BackColor = SystemColors.Control;
            txtInfo_2.BorderStyle = BorderStyle.None;
            txtInfo_2.Location = new Point(68, 70);
            txtInfo_2.Name = "txtInfo_2";
            txtInfo_2.ReadOnly = true;
            txtInfo_2.Size = new Size(161, 13);
            txtInfo_2.TabIndex = 6;
            // 
            // txtInfo_3
            // 
            txtInfo_3.BackColor = SystemColors.Control;
            txtInfo_3.BorderStyle = BorderStyle.None;
            txtInfo_3.Font = new Font("Courier New", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            txtInfo_3.Location = new Point(68, 102);
            txtInfo_3.Name = "txtInfo_3";
            txtInfo_3.ReadOnly = true;
            txtInfo_3.Size = new Size(161, 13);
            txtInfo_3.TabIndex = 5;
            // 
            // Command_1
            // 
            Command_1.Enabled = false;
            Command_1.Location = new Point(164, 126);
            Command_1.Name = "Command_1";
            Command_1.Size = new Size(65, 21);
            Command_1.TabIndex = 4;
            Command_1.Text = "IP Ban";
            Command_1.UseVisualStyleBackColor = true;
            Command_1.Visible = false;
            Command_1.Click += Command_1_Click;
            // 
            // Command_2
            // 
            Command_2.Enabled = false;
            Command_2.Location = new Point(84, 126);
            Command_2.Name = "Command_2";
            Command_2.Size = new Size(65, 21);
            Command_2.TabIndex = 3;
            Command_2.Text = "SID Ban";
            Command_2.UseVisualStyleBackColor = true;
            Command_2.Visible = false;
            Command_2.Click += Command_2_Click;
            // 
            // txtInfo_4
            // 
            txtInfo_4.BackColor = SystemColors.Control;
            txtInfo_4.BorderStyle = BorderStyle.None;
            txtInfo_4.Location = new Point(68, 86);
            txtInfo_4.Name = "txtInfo_4";
            txtInfo_4.ReadOnly = true;
            txtInfo_4.Size = new Size(161, 13);
            txtInfo_4.TabIndex = 2;
            // 
            // Label1
            // 
            Label1.Location = new Point(4, 54);
            Label1.Name = "Label1";
            Label1.Size = new Size(93, 13);
            Label1.TabIndex = 16;
            Label1.Text = "Status:";
            // 
            // Label2
            // 
            Label2.Location = new Point(4, 86);
            Label2.Name = "Label2";
            Label2.Size = new Size(93, 13);
            Label2.TabIndex = 15;
            Label2.Text = "IP Address:";
            // 
            // Label3
            // 
            Label3.Location = new Point(4, 102);
            Label3.Name = "Label3";
            Label3.Size = new Size(73, 17);
            Label3.TabIndex = 14;
            Label3.Text = "Station ID:";
            // 
            // Label4
            // 
            Label4.Location = new Point(4, 70);
            Label4.Name = "Label4";
            Label4.Size = new Size(73, 17);
            Label4.TabIndex = 13;
            Label4.Text = "Authority:";
            // 
            // Label5
            // 
            Label5.Location = new Point(4, 38);
            Label5.Name = "Label5";
            Label5.Size = new Size(93, 13);
            Label5.TabIndex = 12;
            Label5.Text = "Username:";
            // 
            // SSTab1_Page1
            // 
            SSTab1_Page1.Controls.Add(cmdRefresh);
            SSTab1_Page1.Controls.Add(cmdUnban);
            SSTab1_Page1.Controls.Add(BanList);
            SSTab1_Page1.Location = new Point(4, 22);
            SSTab1_Page1.Name = "SSTab1_Page1";
            SSTab1_Page1.Size = new Size(233, 151);
            SSTab1_Page1.TabIndex = 1;
            SSTab1_Page1.Text = "Ban List";
            SSTab1_Page1.UseVisualStyleBackColor = true;
            // 
            // cmdRefresh
            // 
            cmdRefresh.Location = new Point(4, 122);
            cmdRefresh.Name = "cmdRefresh";
            cmdRefresh.Size = new Size(105, 25);
            cmdRefresh.TabIndex = 19;
            cmdRefresh.Text = "Refresh";
            cmdRefresh.UseVisualStyleBackColor = true;
            cmdRefresh.Click += cmdRefresh_Click;
            // 
            // cmdUnban
            // 
            cmdUnban.Location = new Point(124, 122);
            cmdUnban.Name = "cmdUnban";
            cmdUnban.Size = new Size(105, 25);
            cmdUnban.TabIndex = 18;
            cmdUnban.Text = "Unban";
            cmdUnban.UseVisualStyleBackColor = true;
            cmdUnban.Click += cmdUnban_Click;
            // 
            // BanList
            // 
            BanList.Columns.AddRange(new ColumnHeader[] { BanList_Col1, BanList_Col2 });
            BanList.FullRowSelect = true;
            BanList.Location = new Point(4, 6);
            BanList.MultiSelect = false;
            BanList.Name = "BanList";
            BanList.Size = new Size(225, 113);
            BanList.TabIndex = 17;
            BanList.UseCompatibleStateImageBehavior = false;
            BanList.View = View.Details;
            // 
            // BanList_Col1
            // 
            BanList_Col1.Name = "BanList_Col1";
            BanList_Col1.Text = "Name/IP";
            BanList_Col1.Width = 224;
            // 
            // BanList_Col2
            // 
            BanList_Col2.Name = "BanList_Col2";
            BanList_Col2.Text = "Ban Type";
            BanList_Col2.Width = 162;
            // 
            // SSTab1_Page2
            // 
            SSTab1_Page2.Controls.Add(cmdALook);
            SSTab1_Page2.Controls.Add(Text2);
            SSTab1_Page2.Controls.Add(NickList);
            SSTab1_Page2.Location = new Point(4, 24);
            SSTab1_Page2.Name = "SSTab1_Page2";
            SSTab1_Page2.Size = new Size(233, 149);
            SSTab1_Page2.TabIndex = 2;
            SSTab1_Page2.Text = "Alias Lookup";
            SSTab1_Page2.UseVisualStyleBackColor = true;
            // 
            // cmdALook
            // 
            cmdALook.Location = new Point(180, 6);
            cmdALook.Name = "cmdALook";
            cmdALook.Size = new Size(49, 21);
            cmdALook.TabIndex = 22;
            cmdALook.Text = "Lookup";
            cmdALook.UseVisualStyleBackColor = true;
            cmdALook.Click += cmdALook_Click;
            // 
            // Text2
            // 
            Text2.Location = new Point(4, 6);
            Text2.Name = "Text2";
            Text2.Size = new Size(169, 20);
            Text2.TabIndex = 21;
            // 
            // NickList
            // 
            NickList.Columns.AddRange(new ColumnHeader[] { NickList_Col1 });
            NickList.FullRowSelect = true;
            NickList.HeaderStyle = ColumnHeaderStyle.None;
            NickList.Location = new Point(4, 34);
            NickList.MultiSelect = false;
            NickList.Name = "NickList";
            NickList.Size = new Size(225, 113);
            NickList.TabIndex = 20;
            NickList.UseCompatibleStateImageBehavior = false;
            NickList.View = View.Details;
            // 
            // NickList_Col1
            // 
            NickList_Col1.Name = "NickList_Col1";
            NickList_Col1.Text = "";
            NickList_Col1.Width = 224;
            // 
            // StatusBar
            // 
            StatusBar.Items.AddRange(new ToolStripItem[] { StatusBar_Panel1 });
            StatusBar.Location = new Point(0, 180);
            StatusBar.Name = "StatusBar";
            StatusBar.Padding = new Padding(1, 0, 3, 0);
            StatusBar.Size = new Size(250, 22);
            StatusBar.TabIndex = 0;
            StatusBar.Text = "StatusBar";
            // 
            // StatusBar_Panel1
            // 
            StatusBar_Panel1.AutoSize = false;
            StatusBar_Panel1.BorderSides = ToolStripStatusLabelBorderSides.Left | ToolStripStatusLabelBorderSides.Top | ToolStripStatusLabelBorderSides.Right | ToolStripStatusLabelBorderSides.Bottom;
            StatusBar_Panel1.BorderStyle = Border3DStyle.SunkenOuter;
            StatusBar_Panel1.Name = "StatusBar_Panel1";
            StatusBar_Panel1.Size = new Size(96, 17);
            StatusBar_Panel1.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // UserInfo
            // 
            AcceptButton = cmdLookup;
            AutoScaleDimensions = new SizeF(6F, 13F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(250, 202);
            Controls.Add(SSTab1);
            Controls.Add(StatusBar);
            Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "UserInfo";
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Text = "User Info";
            Load += UserInfo_Load;
            SSTab1.ResumeLayout(false);
            SSTab1_Page0.ResumeLayout(false);
            SSTab1_Page0.PerformLayout();
            SSTab1_Page1.ResumeLayout(false);
            SSTab1_Page2.ResumeLayout(false);
            SSTab1_Page2.PerformLayout();
            StatusBar.ResumeLayout(false);
            StatusBar.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.TabControl SSTab1;
        private System.Windows.Forms.TabPage SSTab1_Page0;
        private System.Windows.Forms.TextBox Text1;
        private System.Windows.Forms.Button cmdLookup;
        private System.Windows.Forms.Button Command_0;
        private System.Windows.Forms.TextBox txtInfo_0;
        private System.Windows.Forms.TextBox txtInfo_1;
        private System.Windows.Forms.TextBox txtInfo_2;
        private System.Windows.Forms.TextBox txtInfo_3;
        private System.Windows.Forms.Button Command_1;
        private System.Windows.Forms.Button Command_2;
        private System.Windows.Forms.TextBox txtInfo_4;
        private System.Windows.Forms.Label Label1;
        private System.Windows.Forms.Label Label2;
        private System.Windows.Forms.Label Label3;
        private System.Windows.Forms.Label Label4;
        private System.Windows.Forms.Label Label5;
        private System.Windows.Forms.TabPage SSTab1_Page1;
        private System.Windows.Forms.Button cmdRefresh;
        private System.Windows.Forms.Button cmdUnban;
        private System.Windows.Forms.ListView BanList;
        private System.Windows.Forms.ColumnHeader BanList_Col1;
        private System.Windows.Forms.ColumnHeader BanList_Col2;
        private System.Windows.Forms.TabPage SSTab1_Page2;
        private System.Windows.Forms.Button cmdALook;
        private System.Windows.Forms.TextBox Text2;
        private System.Windows.Forms.ListView NickList;
        private System.Windows.Forms.ColumnHeader NickList_Col1;
        private System.Windows.Forms.StatusStrip StatusBar;
        private System.Windows.Forms.ToolStripStatusLabel StatusBar_Panel1;
    }
}
