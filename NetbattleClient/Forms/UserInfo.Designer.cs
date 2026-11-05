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
            this.SSTab1 = new System.Windows.Forms.TabControl();
            this.SSTab1_Page0 = new System.Windows.Forms.TabPage();
            this.Text1 = new System.Windows.Forms.TextBox();
            this.cmdLookup = new System.Windows.Forms.Button();
            this.Command_0 = new System.Windows.Forms.Button();
            this.txtInfo_0 = new System.Windows.Forms.TextBox();
            this.txtInfo_1 = new System.Windows.Forms.TextBox();
            this.txtInfo_2 = new System.Windows.Forms.TextBox();
            this.txtInfo_3 = new System.Windows.Forms.TextBox();
            this.Command_1 = new System.Windows.Forms.Button();
            this.Command_2 = new System.Windows.Forms.Button();
            this.txtInfo_4 = new System.Windows.Forms.TextBox();
            this.Label1 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label4 = new System.Windows.Forms.Label();
            this.Label5 = new System.Windows.Forms.Label();
            this.SSTab1_Page1 = new System.Windows.Forms.TabPage();
            this.cmdRefresh = new System.Windows.Forms.Button();
            this.cmdUnban = new System.Windows.Forms.Button();
            this.BanList = new System.Windows.Forms.ListView();
            this.BanList_Col1 = new System.Windows.Forms.ColumnHeader();
            this.BanList_Col2 = new System.Windows.Forms.ColumnHeader();
            this.SSTab1_Page2 = new System.Windows.Forms.TabPage();
            this.cmdALook = new System.Windows.Forms.Button();
            this.Text2 = new System.Windows.Forms.TextBox();
            this.NickList = new System.Windows.Forms.ListView();
            this.NickList_Col1 = new System.Windows.Forms.ColumnHeader();
            this.StatusBar = new System.Windows.Forms.StatusStrip();
            this.StatusBar_Panel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.SSTab1.SuspendLayout();
            this.SSTab1_Page0.SuspendLayout();
            this.SSTab1_Page1.SuspendLayout();
            this.SSTab1_Page2.SuspendLayout();
            this.StatusBar.SuspendLayout();
            this.SuspendLayout();
            // 
            // SSTab1
            // 
            this.SSTab1.Controls.Add(this.SSTab1_Page0);
            this.SSTab1.Controls.Add(this.SSTab1_Page1);
            this.SSTab1.Controls.Add(this.SSTab1_Page2);
            this.SSTab1.Location = new System.Drawing.Point(4, 4);
            this.SSTab1.Name = "SSTab1";
            this.SSTab1.Size = new System.Drawing.Size(241, 177);
            this.SSTab1.TabIndex = 1;
            // 
            // SSTab1_Page0
            // 
            this.SSTab1_Page0.Controls.Add(this.Text1);
            this.SSTab1_Page0.Controls.Add(this.cmdLookup);
            this.SSTab1_Page0.Controls.Add(this.Command_0);
            this.SSTab1_Page0.Controls.Add(this.txtInfo_0);
            this.SSTab1_Page0.Controls.Add(this.txtInfo_1);
            this.SSTab1_Page0.Controls.Add(this.txtInfo_2);
            this.SSTab1_Page0.Controls.Add(this.txtInfo_3);
            this.SSTab1_Page0.Controls.Add(this.Command_1);
            this.SSTab1_Page0.Controls.Add(this.Command_2);
            this.SSTab1_Page0.Controls.Add(this.txtInfo_4);
            this.SSTab1_Page0.Controls.Add(this.Label1);
            this.SSTab1_Page0.Controls.Add(this.Label2);
            this.SSTab1_Page0.Controls.Add(this.Label3);
            this.SSTab1_Page0.Controls.Add(this.Label4);
            this.SSTab1_Page0.Controls.Add(this.Label5);
            this.SSTab1_Page0.Location = new System.Drawing.Point(4, 22);
            this.SSTab1_Page0.Name = "SSTab1_Page0";
            this.SSTab1_Page0.Size = new System.Drawing.Size(233, 151);
            this.SSTab1_Page0.TabIndex = 0;
            this.SSTab1_Page0.Text = "User Lookup";
            this.SSTab1_Page0.UseVisualStyleBackColor = true;
            // 
            // Text1
            // 
            this.Text1.Location = new System.Drawing.Point(4, 6);
            this.Text1.Name = "Text1";
            this.Text1.Size = new System.Drawing.Size(169, 21);
            this.Text1.TabIndex = 11;
            this.Text1.Text = "";
            // 
            // cmdLookup
            // 
            this.cmdLookup.Location = new System.Drawing.Point(180, 6);
            this.cmdLookup.Name = "cmdLookup";
            this.cmdLookup.Size = new System.Drawing.Size(49, 21);
            this.cmdLookup.TabIndex = 10;
            this.cmdLookup.Text = "Lookup";
            this.cmdLookup.UseVisualStyleBackColor = true;
            // 
            // Command_0
            // 
            this.Command_0.Location = new System.Drawing.Point(4, 126);
            this.Command_0.Name = "Command_0";
            this.Command_0.Size = new System.Drawing.Size(65, 21);
            this.Command_0.TabIndex = 9;
            this.Command_0.Enabled = false;
            this.Command_0.Text = "TempBan";
            this.Command_0.UseVisualStyleBackColor = true;
            // 
            // txtInfo_0
            // 
            this.txtInfo_0.Location = new System.Drawing.Point(68, 38);
            this.txtInfo_0.Name = "txtInfo_0";
            this.txtInfo_0.Size = new System.Drawing.Size(161, 13);
            this.txtInfo_0.TabIndex = 8;
            this.txtInfo_0.BackColor = System.Drawing.SystemColors.Control;
            this.txtInfo_0.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtInfo_0.ReadOnly = true;
            this.txtInfo_0.Text = "[No User Selected]";
            // 
            // txtInfo_1
            // 
            this.txtInfo_1.Location = new System.Drawing.Point(68, 54);
            this.txtInfo_1.Name = "txtInfo_1";
            this.txtInfo_1.Size = new System.Drawing.Size(161, 13);
            this.txtInfo_1.TabIndex = 7;
            this.txtInfo_1.BackColor = System.Drawing.SystemColors.Control;
            this.txtInfo_1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtInfo_1.ReadOnly = true;
            this.txtInfo_1.Text = "";
            // 
            // txtInfo_2
            // 
            this.txtInfo_2.Location = new System.Drawing.Point(68, 70);
            this.txtInfo_2.Name = "txtInfo_2";
            this.txtInfo_2.Size = new System.Drawing.Size(161, 13);
            this.txtInfo_2.TabIndex = 6;
            this.txtInfo_2.BackColor = System.Drawing.SystemColors.Control;
            this.txtInfo_2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtInfo_2.ReadOnly = true;
            this.txtInfo_2.Text = "";
            // 
            // txtInfo_3
            // 
            this.txtInfo_3.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtInfo_3.Location = new System.Drawing.Point(68, 102);
            this.txtInfo_3.Name = "txtInfo_3";
            this.txtInfo_3.Size = new System.Drawing.Size(161, 14);
            this.txtInfo_3.TabIndex = 5;
            this.txtInfo_3.BackColor = System.Drawing.SystemColors.Control;
            this.txtInfo_3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtInfo_3.ReadOnly = true;
            this.txtInfo_3.Text = "";
            // 
            // Command_1
            // 
            this.Command_1.Location = new System.Drawing.Point(164, 126);
            this.Command_1.Name = "Command_1";
            this.Command_1.Size = new System.Drawing.Size(65, 21);
            this.Command_1.TabIndex = 4;
            this.Command_1.Enabled = false;
            this.Command_1.Visible = false;
            this.Command_1.Text = "IP Ban";
            this.Command_1.UseVisualStyleBackColor = true;
            // 
            // Command_2
            // 
            this.Command_2.Location = new System.Drawing.Point(84, 126);
            this.Command_2.Name = "Command_2";
            this.Command_2.Size = new System.Drawing.Size(65, 21);
            this.Command_2.TabIndex = 3;
            this.Command_2.Enabled = false;
            this.Command_2.Visible = false;
            this.Command_2.Text = "SID Ban";
            this.Command_2.UseVisualStyleBackColor = true;
            // 
            // txtInfo_4
            // 
            this.txtInfo_4.Location = new System.Drawing.Point(68, 86);
            this.txtInfo_4.Name = "txtInfo_4";
            this.txtInfo_4.Size = new System.Drawing.Size(161, 13);
            this.txtInfo_4.TabIndex = 2;
            this.txtInfo_4.BackColor = System.Drawing.SystemColors.Control;
            this.txtInfo_4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtInfo_4.ReadOnly = true;
            this.txtInfo_4.Text = "";
            // 
            // Label1
            // 
            this.Label1.Location = new System.Drawing.Point(4, 54);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(93, 13);
            this.Label1.TabIndex = 16;
            this.Label1.Text = "Status:";
            // 
            // Label2
            // 
            this.Label2.Location = new System.Drawing.Point(4, 86);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(93, 13);
            this.Label2.TabIndex = 15;
            this.Label2.Text = "IP Address:";
            // 
            // Label3
            // 
            this.Label3.Location = new System.Drawing.Point(4, 102);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(73, 17);
            this.Label3.TabIndex = 14;
            this.Label3.Text = "Station ID:";
            // 
            // Label4
            // 
            this.Label4.Location = new System.Drawing.Point(4, 70);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(73, 17);
            this.Label4.TabIndex = 13;
            this.Label4.Text = "Authority:";
            // 
            // Label5
            // 
            this.Label5.Location = new System.Drawing.Point(4, 38);
            this.Label5.Name = "Label5";
            this.Label5.Size = new System.Drawing.Size(93, 13);
            this.Label5.TabIndex = 12;
            this.Label5.Text = "Username:";
            // 
            // SSTab1_Page1
            // 
            this.SSTab1_Page1.Controls.Add(this.cmdRefresh);
            this.SSTab1_Page1.Controls.Add(this.cmdUnban);
            this.SSTab1_Page1.Controls.Add(this.BanList);
            this.SSTab1_Page1.Location = new System.Drawing.Point(4, 22);
            this.SSTab1_Page1.Name = "SSTab1_Page1";
            this.SSTab1_Page1.Size = new System.Drawing.Size(233, 151);
            this.SSTab1_Page1.TabIndex = 1;
            this.SSTab1_Page1.Text = "Ban List";
            this.SSTab1_Page1.UseVisualStyleBackColor = true;
            // 
            // cmdRefresh
            // 
            this.cmdRefresh.Location = new System.Drawing.Point(4, 122);
            this.cmdRefresh.Name = "cmdRefresh";
            this.cmdRefresh.Size = new System.Drawing.Size(105, 25);
            this.cmdRefresh.TabIndex = 19;
            this.cmdRefresh.Text = "Refresh";
            this.cmdRefresh.UseVisualStyleBackColor = true;
            // 
            // cmdUnban
            // 
            this.cmdUnban.Location = new System.Drawing.Point(124, 122);
            this.cmdUnban.Name = "cmdUnban";
            this.cmdUnban.Size = new System.Drawing.Size(105, 25);
            this.cmdUnban.TabIndex = 18;
            this.cmdUnban.Text = "Unban";
            this.cmdUnban.UseVisualStyleBackColor = true;
            // 
            // BanList
            // 
            this.BanList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
                        this.BanList_Col1,
                        this.BanList_Col2});
            this.BanList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.BanList.FullRowSelect = true;
            this.BanList.HideSelection = false;
            this.BanList.MultiSelect = false;
            this.BanList.Location = new System.Drawing.Point(4, 6);
            this.BanList.Name = "BanList";
            this.BanList.Size = new System.Drawing.Size(225, 113);
            this.BanList.TabIndex = 17;
            this.BanList.UseCompatibleStateImageBehavior = false;
            this.BanList.View = System.Windows.Forms.View.Details;
            // 
            // BanList_Col1
            // 
            this.BanList_Col1.Text = "Name/IP";
            this.BanList_Col1.Width = 224;
            // 
            // BanList_Col2
            // 
            this.BanList_Col2.Text = "Ban Type";
            this.BanList_Col2.Width = 162;
            // 
            // SSTab1_Page2
            // 
            this.SSTab1_Page2.Controls.Add(this.cmdALook);
            this.SSTab1_Page2.Controls.Add(this.Text2);
            this.SSTab1_Page2.Controls.Add(this.NickList);
            this.SSTab1_Page2.Location = new System.Drawing.Point(4, 22);
            this.SSTab1_Page2.Name = "SSTab1_Page2";
            this.SSTab1_Page2.Size = new System.Drawing.Size(233, 151);
            this.SSTab1_Page2.TabIndex = 2;
            this.SSTab1_Page2.Text = "Alias Lookup";
            this.SSTab1_Page2.UseVisualStyleBackColor = true;
            // 
            // cmdALook
            // 
            this.cmdALook.Location = new System.Drawing.Point(180, 6);
            this.cmdALook.Name = "cmdALook";
            this.cmdALook.Size = new System.Drawing.Size(49, 21);
            this.cmdALook.TabIndex = 22;
            this.cmdALook.Text = "Lookup";
            this.cmdALook.UseVisualStyleBackColor = true;
            // 
            // Text2
            // 
            this.Text2.Location = new System.Drawing.Point(4, 6);
            this.Text2.Name = "Text2";
            this.Text2.Size = new System.Drawing.Size(169, 21);
            this.Text2.TabIndex = 21;
            this.Text2.Text = "";
            // 
            // NickList
            // 
            this.NickList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
                        this.NickList_Col1});
            this.NickList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.NickList.FullRowSelect = true;
            this.NickList.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.NickList.HideSelection = false;
            this.NickList.MultiSelect = false;
            this.NickList.Location = new System.Drawing.Point(4, 34);
            this.NickList.Name = "NickList";
            this.NickList.Size = new System.Drawing.Size(225, 113);
            this.NickList.TabIndex = 20;
            this.NickList.UseCompatibleStateImageBehavior = false;
            this.NickList.View = System.Windows.Forms.View.Details;
            // 
            // NickList_Col1
            // 
            this.NickList_Col1.Text = "";
            this.NickList_Col1.Width = 224;
            // 
            // StatusBar
            // 
            this.StatusBar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
                        this.StatusBar_Panel1});
            this.StatusBar.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.StatusBar.Location = new System.Drawing.Point(0, 185);
            this.StatusBar.Name = "StatusBar";
            this.StatusBar.Size = new System.Drawing.Size(250, 17);
            this.StatusBar.TabIndex = 0;
            this.StatusBar.Text = "StatusBar";
            // 
            // StatusBar_Panel1
            // 
            this.StatusBar_Panel1.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.All;
            this.StatusBar_Panel1.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter;
            this.StatusBar_Panel1.Name = "StatusBar_Panel1";
            this.StatusBar_Panel1.AutoSize = false;
            this.StatusBar_Panel1.Size = new System.Drawing.Size(96, 17);
            this.StatusBar_Panel1.Text = "";
            this.StatusBar_Panel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // UserInfo
            // 
            this.AcceptButton = this.cmdLookup;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(250, 202);
            this.Controls.Add(this.SSTab1);
            this.Controls.Add(this.StatusBar);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "UserInfo";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "User Info";
            this.SSTab1.ResumeLayout(false);
            this.SSTab1.PerformLayout();
            this.SSTab1_Page0.ResumeLayout(false);
            this.SSTab1_Page0.PerformLayout();
            this.SSTab1_Page1.ResumeLayout(false);
            this.SSTab1_Page1.PerformLayout();
            this.SSTab1_Page2.ResumeLayout(false);
            this.SSTab1_Page2.PerformLayout();
            this.StatusBar.ResumeLayout(false);
            this.StatusBar.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
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
