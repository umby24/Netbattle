namespace Netbattle.Forms
{
    partial class ServerDataManager
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
            this.components = new System.ComponentModel.Container();
            this.Picture1 = new System.Windows.Forms.Panel();
            this.ProgBar = new System.Windows.Forms.ProgressBar();
            this.Label6 = new System.Windows.Forms.Label();
            this.Command1 = new System.Windows.Forms.Button();
            this.SSTab1 = new System.Windows.Forms.TabControl();
            this.SSTab1_Page0 = new System.Windows.Forms.TabPage();
            this.AuthSet_3 = new System.Windows.Forms.RadioButton();
            this.NewPW = new System.Windows.Forms.TextBox();
            this.ChangePW = new System.Windows.Forms.Button();
            this.DeleteUser = new System.Windows.Forms.Button();
            this.AuthChange = new System.Windows.Forms.Button();
            this.AuthSet_0 = new System.Windows.Forms.RadioButton();
            this.AuthSet_1 = new System.Windows.Forms.RadioButton();
            this.AuthSet_2 = new System.Windows.Forms.RadioButton();
            this.NewUser = new System.Windows.Forms.TextBox();
            this.Command9 = new System.Windows.Forms.Button();
            this.txtUserJump = new System.Windows.Forms.TextBox();
            this.UserList = new System.Windows.Forms.ListView();
            this.UserList_Col1 = new System.Windows.Forms.ColumnHeader();
            this.UserList_Col2 = new System.Windows.Forms.ColumnHeader();
            this.Label1 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.SSTab1_Page1 = new System.Windows.Forms.TabPage();
            this.Command7 = new System.Windows.Forms.Button();
            this.List4 = new System.Windows.Forms.ListBox();
            this.Command6 = new System.Windows.Forms.Button();
            this.Text3 = new System.Windows.Forms.TextBox();
            this.SSTab1_Page2 = new System.Windows.Forms.TabPage();
            this.cmdChange2 = new System.Windows.Forms.Button();
            this.txtBan2 = new System.Windows.Forms.TextBox();
            this.Text1 = new System.Windows.Forms.TextBox();
            this.Command3 = new System.Windows.Forms.Button();
            this.List2 = new System.Windows.Forms.ListBox();
            this.Command2 = new System.Windows.Forms.Button();
            this.Label3 = new System.Windows.Forms.Label();
            this.Line3 = new System.Windows.Forms.Label();
            this.SSTab1_Page3 = new System.Windows.Forms.TabPage();
            this.cmdChange1 = new System.Windows.Forms.Button();
            this.txtBan1 = new System.Windows.Forms.TextBox();
            this.Text5 = new System.Windows.Forms.TextBox();
            this.Command11 = new System.Windows.Forms.Button();
            this.List1 = new System.Windows.Forms.ListBox();
            this.Command10 = new System.Windows.Forms.Button();
            this.Label5 = new System.Windows.Forms.Label();
            this.Line1 = new System.Windows.Forms.Label();
            this.SSTab1_Page4 = new System.Windows.Forms.TabPage();
            this.cmdChange3 = new System.Windows.Forms.Button();
            this.txtBan3 = new System.Windows.Forms.TextBox();
            this.Command5 = new System.Windows.Forms.Button();
            this.List3 = new System.Windows.Forms.ListBox();
            this.Command4 = new System.Windows.Forms.Button();
            this.Text2 = new System.Windows.Forms.TextBox();
            this.Line2 = new System.Windows.Forms.Label();
            this.Label4 = new System.Windows.Forms.Label();
            this.tmrCloser = new System.Windows.Forms.Timer(this.components);
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.Picture1.SuspendLayout();
            this.SSTab1.SuspendLayout();
            this.SSTab1_Page0.SuspendLayout();
            this.SSTab1_Page1.SuspendLayout();
            this.SSTab1_Page2.SuspendLayout();
            this.SSTab1_Page3.SuspendLayout();
            this.SSTab1_Page4.SuspendLayout();
            this.SuspendLayout();
            // 
            // Picture1
            // 
            this.Picture1.Controls.Add(this.ProgBar);
            this.Picture1.Controls.Add(this.Label6);
            this.Picture1.Location = new System.Drawing.Point(16, 280);
            this.Picture1.Name = "Picture1";
            this.Picture1.Size = new System.Drawing.Size(321, 81);
            this.Picture1.TabIndex = 1;
            // 
            // ProgBar
            // 
            this.ProgBar.Location = new System.Drawing.Point(8, 40);
            this.ProgBar.Name = "ProgBar";
            this.ProgBar.Size = new System.Drawing.Size(305, 17);
            // 
            // Label6
            // 
            this.Label6.Location = new System.Drawing.Point(8, 24);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(305, 17);
            this.Label6.TabIndex = 2;
            this.Label6.Text = "Populating User List...";
            this.Label6.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Command1
            // 
            this.Command1.Location = new System.Drawing.Point(256, 288);
            this.Command1.Name = "Command1";
            this.Command1.Size = new System.Drawing.Size(81, 25);
            this.Command1.TabIndex = 0;
            this.Command1.Text = "&Ok";
            this.Command1.UseVisualStyleBackColor = true;
            // 
            // SSTab1
            // 
            this.SSTab1.Controls.Add(this.SSTab1_Page0);
            this.SSTab1.Controls.Add(this.SSTab1_Page1);
            this.SSTab1.Controls.Add(this.SSTab1_Page2);
            this.SSTab1.Controls.Add(this.SSTab1_Page3);
            this.SSTab1.Controls.Add(this.SSTab1_Page4);
            this.SSTab1.Location = new System.Drawing.Point(8, 8);
            this.SSTab1.Name = "SSTab1";
            this.SSTab1.Size = new System.Drawing.Size(337, 273);
            this.SSTab1.TabIndex = 3;
            this.SSTab1.Enabled = false;
            // 
            // SSTab1_Page0
            // 
            this.SSTab1_Page0.Controls.Add(this.AuthSet_3);
            this.SSTab1_Page0.Controls.Add(this.NewPW);
            this.SSTab1_Page0.Controls.Add(this.ChangePW);
            this.SSTab1_Page0.Controls.Add(this.DeleteUser);
            this.SSTab1_Page0.Controls.Add(this.AuthChange);
            this.SSTab1_Page0.Controls.Add(this.AuthSet_0);
            this.SSTab1_Page0.Controls.Add(this.AuthSet_1);
            this.SSTab1_Page0.Controls.Add(this.AuthSet_2);
            this.SSTab1_Page0.Controls.Add(this.NewUser);
            this.SSTab1_Page0.Controls.Add(this.Command9);
            this.SSTab1_Page0.Controls.Add(this.txtUserJump);
            this.SSTab1_Page0.Controls.Add(this.UserList);
            this.SSTab1_Page0.Controls.Add(this.Label1);
            this.SSTab1_Page0.Controls.Add(this.Label2);
            this.SSTab1_Page0.Location = new System.Drawing.Point(4, 22);
            this.SSTab1_Page0.Name = "SSTab1_Page0";
            this.SSTab1_Page0.Size = new System.Drawing.Size(329, 247);
            this.SSTab1_Page0.TabIndex = 0;
            this.SSTab1_Page0.Text = "User List";
            this.SSTab1_Page0.UseVisualStyleBackColor = true;
            // 
            // AuthSet_3
            // 
            this.AuthSet_3.Location = new System.Drawing.Point(4, 202);
            this.AuthSet_3.Name = "AuthSet_3";
            this.AuthSet_3.Size = new System.Drawing.Size(89, 17);
            this.AuthSet_3.TabIndex = 42;
            this.AuthSet_3.Text = "&Mega-Admin";
            this.AuthSet_3.UseVisualStyleBackColor = true;
            // 
            // NewPW
            // 
            this.NewPW.Location = new System.Drawing.Point(164, 194);
            this.NewPW.Name = "NewPW";
            this.NewPW.Size = new System.Drawing.Size(161, 19);
            this.NewPW.TabIndex = 36;
            this.NewPW.PasswordChar = '*';
            this.NewPW.Text = "";
            // 
            // ChangePW
            // 
            this.ChangePW.Location = new System.Drawing.Point(168, 218);
            this.ChangePW.Name = "ChangePW";
            this.ChangePW.Size = new System.Drawing.Size(73, 25);
            this.ChangePW.TabIndex = 35;
            this.ChangePW.Text = "&Password";
            this.ChangePW.UseVisualStyleBackColor = true;
            // 
            // DeleteUser
            // 
            this.DeleteUser.Location = new System.Drawing.Point(248, 218);
            this.DeleteUser.Name = "DeleteUser";
            this.DeleteUser.Size = new System.Drawing.Size(73, 25);
            this.DeleteUser.TabIndex = 34;
            this.DeleteUser.Text = "&Delete";
            this.DeleteUser.UseVisualStyleBackColor = true;
            // 
            // AuthChange
            // 
            this.AuthChange.Location = new System.Drawing.Point(8, 218);
            this.AuthChange.Name = "AuthChange";
            this.AuthChange.Size = new System.Drawing.Size(73, 25);
            this.AuthChange.TabIndex = 32;
            this.toolTip.SetToolTip(this.AuthChange, "Change user authority");
            this.AuthChange.Text = "&Authority";
            this.AuthChange.UseVisualStyleBackColor = true;
            // 
            // AuthSet_0
            // 
            this.AuthSet_0.Location = new System.Drawing.Point(4, 154);
            this.AuthSet_0.Name = "AuthSet_0";
            this.AuthSet_0.Size = new System.Drawing.Size(49, 17);
            this.AuthSet_0.TabIndex = 31;
            this.AuthSet_0.Checked = true;
            this.AuthSet_0.Text = "&User";
            this.AuthSet_0.UseVisualStyleBackColor = true;
            // 
            // AuthSet_1
            // 
            this.AuthSet_1.Location = new System.Drawing.Point(4, 170);
            this.AuthSet_1.Name = "AuthSet_1";
            this.AuthSet_1.Size = new System.Drawing.Size(49, 17);
            this.AuthSet_1.TabIndex = 30;
            this.AuthSet_1.Text = "&Mod";
            this.AuthSet_1.UseVisualStyleBackColor = true;
            // 
            // AuthSet_2
            // 
            this.AuthSet_2.Location = new System.Drawing.Point(4, 186);
            this.AuthSet_2.Name = "AuthSet_2";
            this.AuthSet_2.Size = new System.Drawing.Size(57, 17);
            this.AuthSet_2.TabIndex = 29;
            this.AuthSet_2.Text = "&Admin";
            this.AuthSet_2.UseVisualStyleBackColor = true;
            // 
            // NewUser
            // 
            this.NewUser.Location = new System.Drawing.Point(164, 170);
            this.NewUser.Name = "NewUser";
            this.NewUser.Size = new System.Drawing.Size(161, 19);
            this.NewUser.TabIndex = 28;
            this.NewUser.Text = "";
            // 
            // Command9
            // 
            this.Command9.Location = new System.Drawing.Point(88, 218);
            this.Command9.Name = "Command9";
            this.Command9.Size = new System.Drawing.Size(73, 25);
            this.Command9.TabIndex = 27;
            this.Command9.Text = "SID &Ban";
            this.Command9.UseVisualStyleBackColor = true;
            // 
            // txtUserJump
            // 
            this.txtUserJump.Location = new System.Drawing.Point(4, 2);
            this.txtUserJump.Name = "txtUserJump";
            this.txtUserJump.Size = new System.Drawing.Size(321, 19);
            this.txtUserJump.TabIndex = 26;
            this.txtUserJump.Text = "";
            // 
            // UserList
            // 
            this.UserList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
                        this.UserList_Col1,
                        this.UserList_Col2});
            this.UserList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.UserList.FullRowSelect = true;
            this.UserList.HideSelection = false;
            this.UserList.MultiSelect = false;
            this.UserList.Location = new System.Drawing.Point(4, 26);
            this.UserList.Name = "UserList";
            this.UserList.Size = new System.Drawing.Size(321, 129);
            this.UserList.TabIndex = 33;
            this.UserList.UseCompatibleStateImageBehavior = false;
            this.UserList.View = System.Windows.Forms.View.Details;
            // 
            // UserList_Col1
            // 
            this.UserList_Col1.Text = "Player Name";
            this.UserList_Col1.Width = 340;
            // 
            // UserList_Col2
            // 
            this.UserList_Col2.Text = "Authority";
            this.UserList_Col2.Width = 141;
            // 
            // Label1
            // 
            this.Label1.Location = new System.Drawing.Point(100, 194);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(57, 17);
            this.Label1.TabIndex = 41;
            this.Label1.BackColor = System.Drawing.Color.Transparent;
            this.Label1.Text = "Password";
            this.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // Label2
            // 
            this.Label2.Location = new System.Drawing.Point(100, 170);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(57, 17);
            this.Label2.TabIndex = 40;
            this.Label2.BackColor = System.Drawing.Color.Transparent;
            this.Label2.Text = "User Name";
            this.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // SSTab1_Page1
            // 
            this.SSTab1_Page1.Controls.Add(this.Command7);
            this.SSTab1_Page1.Controls.Add(this.List4);
            this.SSTab1_Page1.Controls.Add(this.Command6);
            this.SSTab1_Page1.Controls.Add(this.Text3);
            this.SSTab1_Page1.Location = new System.Drawing.Point(4, 22);
            this.SSTab1_Page1.Name = "SSTab1_Page1";
            this.SSTab1_Page1.Size = new System.Drawing.Size(329, 247);
            this.SSTab1_Page1.TabIndex = 1;
            this.SSTab1_Page1.Text = "Word Filter";
            this.SSTab1_Page1.UseVisualStyleBackColor = true;
            // 
            // Command7
            // 
            this.Command7.Location = new System.Drawing.Point(252, 218);
            this.Command7.Name = "Command7";
            this.Command7.Size = new System.Drawing.Size(73, 25);
            this.Command7.TabIndex = 7;
            this.Command7.Text = "&Remove";
            this.Command7.UseVisualStyleBackColor = true;
            // 
            // List4
            // 
            this.List4.Location = new System.Drawing.Point(4, 2);
            this.List4.Name = "List4";
            this.List4.Size = new System.Drawing.Size(321, 199);
            this.List4.TabIndex = 6;
            this.List4.FormattingEnabled = true;
            // 
            // Command6
            // 
            this.Command6.Location = new System.Drawing.Point(172, 218);
            this.Command6.Name = "Command6";
            this.Command6.Size = new System.Drawing.Size(73, 25);
            this.Command6.TabIndex = 5;
            this.Command6.Text = "&Add";
            this.Command6.UseVisualStyleBackColor = true;
            // 
            // Text3
            // 
            this.Text3.Location = new System.Drawing.Point(4, 221);
            this.Text3.Name = "Text3";
            this.Text3.Size = new System.Drawing.Size(161, 19);
            this.Text3.TabIndex = 4;
            this.Text3.Text = "";
            // 
            // SSTab1_Page2
            // 
            this.SSTab1_Page2.Controls.Add(this.cmdChange2);
            this.SSTab1_Page2.Controls.Add(this.txtBan2);
            this.SSTab1_Page2.Controls.Add(this.Text1);
            this.SSTab1_Page2.Controls.Add(this.Command3);
            this.SSTab1_Page2.Controls.Add(this.List2);
            this.SSTab1_Page2.Controls.Add(this.Command2);
            this.SSTab1_Page2.Controls.Add(this.Label3);
            this.SSTab1_Page2.Controls.Add(this.Line3);
            this.SSTab1_Page2.Location = new System.Drawing.Point(4, 22);
            this.SSTab1_Page2.Name = "SSTab1_Page2";
            this.SSTab1_Page2.Size = new System.Drawing.Size(329, 247);
            this.SSTab1_Page2.TabIndex = 2;
            this.SSTab1_Page2.Text = "IP Ban";
            this.SSTab1_Page2.UseVisualStyleBackColor = true;
            // 
            // cmdChange2
            // 
            this.cmdChange2.Location = new System.Drawing.Point(252, 178);
            this.cmdChange2.Name = "cmdChange2";
            this.cmdChange2.Size = new System.Drawing.Size(73, 25);
            this.cmdChange2.TabIndex = 13;
            this.cmdChange2.Text = "Change";
            this.cmdChange2.UseVisualStyleBackColor = true;
            // 
            // txtBan2
            // 
            this.txtBan2.Location = new System.Drawing.Point(4, 181);
            this.txtBan2.Name = "txtBan2";
            this.txtBan2.Size = new System.Drawing.Size(241, 19);
            this.txtBan2.TabIndex = 12;
            this.txtBan2.Text = "";
            // 
            // Text1
            // 
            this.Text1.Location = new System.Drawing.Point(4, 221);
            this.Text1.Name = "Text1";
            this.Text1.Size = new System.Drawing.Size(161, 19);
            this.Text1.TabIndex = 11;
            this.Text1.Text = "";
            // 
            // Command3
            // 
            this.Command3.Location = new System.Drawing.Point(172, 218);
            this.Command3.Name = "Command3";
            this.Command3.Size = new System.Drawing.Size(73, 25);
            this.Command3.TabIndex = 10;
            this.Command3.Text = "&Manual Ban";
            this.Command3.UseVisualStyleBackColor = true;
            // 
            // List2
            // 
            this.List2.Location = new System.Drawing.Point(4, 2);
            this.List2.Name = "List2";
            this.List2.Size = new System.Drawing.Size(321, 160);
            this.List2.TabIndex = 9;
            this.List2.FormattingEnabled = true;
            // 
            // Command2
            // 
            this.Command2.Location = new System.Drawing.Point(252, 218);
            this.Command2.Name = "Command2";
            this.Command2.Size = new System.Drawing.Size(73, 25);
            this.Command2.TabIndex = 8;
            this.Command2.Text = "&Unban";
            this.Command2.UseVisualStyleBackColor = true;
            // 
            // Label3
            // 
            this.Label3.Location = new System.Drawing.Point(4, 166);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(129, 17);
            this.Label3.TabIndex = 37;
            this.Label3.Text = "Ban Message:";
            // 
            // Line3
            // 
            this.Line3.Location = new System.Drawing.Point(0, 210);
            this.Line3.Name = "Line3";
            this.Line3.Size = new System.Drawing.Size(336, 1);
            this.Line3.BackColor = System.Drawing.SystemColors.WindowText;
            // 
            // SSTab1_Page3
            // 
            this.SSTab1_Page3.Controls.Add(this.cmdChange1);
            this.SSTab1_Page3.Controls.Add(this.txtBan1);
            this.SSTab1_Page3.Controls.Add(this.Text5);
            this.SSTab1_Page3.Controls.Add(this.Command11);
            this.SSTab1_Page3.Controls.Add(this.List1);
            this.SSTab1_Page3.Controls.Add(this.Command10);
            this.SSTab1_Page3.Controls.Add(this.Label5);
            this.SSTab1_Page3.Controls.Add(this.Line1);
            this.SSTab1_Page3.Location = new System.Drawing.Point(4, 22);
            this.SSTab1_Page3.Name = "SSTab1_Page3";
            this.SSTab1_Page3.Size = new System.Drawing.Size(329, 247);
            this.SSTab1_Page3.TabIndex = 3;
            this.SSTab1_Page3.Text = "SID Ban";
            this.SSTab1_Page3.UseVisualStyleBackColor = true;
            // 
            // cmdChange1
            // 
            this.cmdChange1.Location = new System.Drawing.Point(252, 178);
            this.cmdChange1.Name = "cmdChange1";
            this.cmdChange1.Size = new System.Drawing.Size(73, 25);
            this.cmdChange1.TabIndex = 25;
            this.cmdChange1.Text = "Change";
            this.cmdChange1.UseVisualStyleBackColor = true;
            // 
            // txtBan1
            // 
            this.txtBan1.Location = new System.Drawing.Point(4, 181);
            this.txtBan1.Name = "txtBan1";
            this.txtBan1.Size = new System.Drawing.Size(241, 19);
            this.txtBan1.TabIndex = 24;
            this.txtBan1.Text = "";
            // 
            // Text5
            // 
            this.Text5.Location = new System.Drawing.Point(4, 221);
            this.Text5.Name = "Text5";
            this.Text5.Size = new System.Drawing.Size(161, 19);
            this.Text5.TabIndex = 23;
            this.Text5.Text = "";
            // 
            // Command11
            // 
            this.Command11.Location = new System.Drawing.Point(172, 218);
            this.Command11.Name = "Command11";
            this.Command11.Size = new System.Drawing.Size(73, 25);
            this.Command11.TabIndex = 22;
            this.Command11.Text = "&Ban";
            this.Command11.UseVisualStyleBackColor = true;
            // 
            // List1
            // 
            this.List1.Location = new System.Drawing.Point(4, 2);
            this.List1.Name = "List1";
            this.List1.Size = new System.Drawing.Size(321, 160);
            this.List1.TabIndex = 21;
            this.List1.FormattingEnabled = true;
            // 
            // Command10
            // 
            this.Command10.Location = new System.Drawing.Point(252, 218);
            this.Command10.Name = "Command10";
            this.Command10.Size = new System.Drawing.Size(73, 25);
            this.Command10.TabIndex = 20;
            this.Command10.Text = "&Unban";
            this.Command10.UseVisualStyleBackColor = true;
            // 
            // Label5
            // 
            this.Label5.Location = new System.Drawing.Point(4, 166);
            this.Label5.Name = "Label5";
            this.Label5.Size = new System.Drawing.Size(129, 17);
            this.Label5.TabIndex = 39;
            this.Label5.Text = "Ban Message:";
            // 
            // Line1
            // 
            this.Line1.Location = new System.Drawing.Point(0, 210);
            this.Line1.Name = "Line1";
            this.Line1.Size = new System.Drawing.Size(336, 1);
            this.Line1.BackColor = System.Drawing.SystemColors.WindowText;
            // 
            // SSTab1_Page4
            // 
            this.SSTab1_Page4.Controls.Add(this.cmdChange3);
            this.SSTab1_Page4.Controls.Add(this.txtBan3);
            this.SSTab1_Page4.Controls.Add(this.Command5);
            this.SSTab1_Page4.Controls.Add(this.List3);
            this.SSTab1_Page4.Controls.Add(this.Command4);
            this.SSTab1_Page4.Controls.Add(this.Text2);
            this.SSTab1_Page4.Controls.Add(this.Line2);
            this.SSTab1_Page4.Controls.Add(this.Label4);
            this.SSTab1_Page4.Location = new System.Drawing.Point(4, 22);
            this.SSTab1_Page4.Name = "SSTab1_Page4";
            this.SSTab1_Page4.Size = new System.Drawing.Size(329, 247);
            this.SSTab1_Page4.TabIndex = 4;
            this.SSTab1_Page4.Text = "ISP Ban";
            this.SSTab1_Page4.UseVisualStyleBackColor = true;
            // 
            // cmdChange3
            // 
            this.cmdChange3.Location = new System.Drawing.Point(252, 178);
            this.cmdChange3.Name = "cmdChange3";
            this.cmdChange3.Size = new System.Drawing.Size(73, 25);
            this.cmdChange3.TabIndex = 19;
            this.cmdChange3.Text = "Change";
            this.cmdChange3.UseVisualStyleBackColor = true;
            // 
            // txtBan3
            // 
            this.txtBan3.Location = new System.Drawing.Point(4, 181);
            this.txtBan3.Name = "txtBan3";
            this.txtBan3.Size = new System.Drawing.Size(241, 19);
            this.txtBan3.TabIndex = 18;
            this.txtBan3.Text = "";
            // 
            // Command5
            // 
            this.Command5.Location = new System.Drawing.Point(252, 218);
            this.Command5.Name = "Command5";
            this.Command5.Size = new System.Drawing.Size(73, 25);
            this.Command5.TabIndex = 17;
            this.Command5.Text = "&Unban";
            this.Command5.UseVisualStyleBackColor = true;
            // 
            // List3
            // 
            this.List3.Location = new System.Drawing.Point(4, 2);
            this.List3.Name = "List3";
            this.List3.Size = new System.Drawing.Size(321, 160);
            this.List3.TabIndex = 16;
            this.List3.FormattingEnabled = true;
            // 
            // Command4
            // 
            this.Command4.Location = new System.Drawing.Point(172, 218);
            this.Command4.Name = "Command4";
            this.Command4.Size = new System.Drawing.Size(73, 25);
            this.Command4.TabIndex = 15;
            this.Command4.Text = "&Manual Ban";
            this.Command4.UseVisualStyleBackColor = true;
            // 
            // Text2
            // 
            this.Text2.Location = new System.Drawing.Point(4, 221);
            this.Text2.Name = "Text2";
            this.Text2.Size = new System.Drawing.Size(161, 19);
            this.Text2.TabIndex = 14;
            this.Text2.Text = "";
            // 
            // Line2
            // 
            this.Line2.Location = new System.Drawing.Point(0, 210);
            this.Line2.Name = "Line2";
            this.Line2.Size = new System.Drawing.Size(336, 1);
            this.Line2.BackColor = System.Drawing.SystemColors.WindowText;
            // 
            // Label4
            // 
            this.Label4.Location = new System.Drawing.Point(4, 166);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(129, 17);
            this.Label4.TabIndex = 38;
            this.Label4.Text = "Ban Message:";
            // 
            // tmrCloser
            // 
            this.tmrCloser.Interval = 1;
            // 
            // ServerDataManager
            // 
            this.AcceptButton = this.Command1;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.Command1;
            this.ClientSize = new System.Drawing.Size(351, 319);
            this.Controls.Add(this.Picture1);
            this.Controls.Add(this.Command1);
            this.Controls.Add(this.SSTab1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ServerDataManager";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Server Data Manager";
            this.Picture1.ResumeLayout(false);
            this.Picture1.PerformLayout();
            this.SSTab1.ResumeLayout(false);
            this.SSTab1.PerformLayout();
            this.SSTab1_Page0.ResumeLayout(false);
            this.SSTab1_Page0.PerformLayout();
            this.SSTab1_Page1.ResumeLayout(false);
            this.SSTab1_Page1.PerformLayout();
            this.SSTab1_Page2.ResumeLayout(false);
            this.SSTab1_Page2.PerformLayout();
            this.SSTab1_Page3.ResumeLayout(false);
            this.SSTab1_Page3.PerformLayout();
            this.SSTab1_Page4.ResumeLayout(false);
            this.SSTab1_Page4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Panel Picture1;
        private System.Windows.Forms.ProgressBar ProgBar;
        private System.Windows.Forms.Label Label6;
        private System.Windows.Forms.Button Command1;
        private System.Windows.Forms.TabControl SSTab1;
        private System.Windows.Forms.TabPage SSTab1_Page0;
        private System.Windows.Forms.RadioButton AuthSet_3;
        private System.Windows.Forms.TextBox NewPW;
        private System.Windows.Forms.Button ChangePW;
        private System.Windows.Forms.Button DeleteUser;
        private System.Windows.Forms.Button AuthChange;
        private System.Windows.Forms.RadioButton AuthSet_0;
        private System.Windows.Forms.RadioButton AuthSet_1;
        private System.Windows.Forms.RadioButton AuthSet_2;
        private System.Windows.Forms.TextBox NewUser;
        private System.Windows.Forms.Button Command9;
        private System.Windows.Forms.TextBox txtUserJump;
        private System.Windows.Forms.ListView UserList;
        private System.Windows.Forms.ColumnHeader UserList_Col1;
        private System.Windows.Forms.ColumnHeader UserList_Col2;
        private System.Windows.Forms.Label Label1;
        private System.Windows.Forms.Label Label2;
        private System.Windows.Forms.TabPage SSTab1_Page1;
        private System.Windows.Forms.Button Command7;
        private System.Windows.Forms.ListBox List4;
        private System.Windows.Forms.Button Command6;
        private System.Windows.Forms.TextBox Text3;
        private System.Windows.Forms.TabPage SSTab1_Page2;
        private System.Windows.Forms.Button cmdChange2;
        private System.Windows.Forms.TextBox txtBan2;
        private System.Windows.Forms.TextBox Text1;
        private System.Windows.Forms.Button Command3;
        private System.Windows.Forms.ListBox List2;
        private System.Windows.Forms.Button Command2;
        private System.Windows.Forms.Label Label3;
        private System.Windows.Forms.Label Line3;
        private System.Windows.Forms.TabPage SSTab1_Page3;
        private System.Windows.Forms.Button cmdChange1;
        private System.Windows.Forms.TextBox txtBan1;
        private System.Windows.Forms.TextBox Text5;
        private System.Windows.Forms.Button Command11;
        private System.Windows.Forms.ListBox List1;
        private System.Windows.Forms.Button Command10;
        private System.Windows.Forms.Label Label5;
        private System.Windows.Forms.Label Line1;
        private System.Windows.Forms.TabPage SSTab1_Page4;
        private System.Windows.Forms.Button cmdChange3;
        private System.Windows.Forms.TextBox txtBan3;
        private System.Windows.Forms.Button Command5;
        private System.Windows.Forms.ListBox List3;
        private System.Windows.Forms.Button Command4;
        private System.Windows.Forms.TextBox Text2;
        private System.Windows.Forms.Label Line2;
        private System.Windows.Forms.Label Label4;
        private System.Windows.Forms.Timer tmrCloser;
        private System.Windows.Forms.ToolTip toolTip;
    }
}
