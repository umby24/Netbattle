namespace Netbattle.Forms
{
    partial class ServerScript
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
            this.Placeholder = new System.Windows.Forms.CheckBox();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.cmdOK = new System.Windows.Forms.Button();
            this.SSTab1 = new System.Windows.Forms.TabControl();
            this.SSTab1_Page0 = new System.Windows.Forms.TabPage();
            this.Check1 = new System.Windows.Forms.CheckBox();
            this.txtScript = new System.Windows.Forms.RichTextBox();
            this.lblLineNum = new System.Windows.Forms.Label();
            this.SSTab1_Page1 = new System.Windows.Forms.TabPage();
            this.txtMsg_9 = new System.Windows.Forms.TextBox();
            this.txtMsg_8 = new System.Windows.Forms.TextBox();
            this.txtMsg_7 = new System.Windows.Forms.TextBox();
            this.txtMsg_6 = new System.Windows.Forms.TextBox();
            this.txtMsg_5 = new System.Windows.Forms.TextBox();
            this.txtMsg_4 = new System.Windows.Forms.TextBox();
            this.txtMsg_3 = new System.Windows.Forms.TextBox();
            this.txtMsg_2 = new System.Windows.Forms.TextBox();
            this.txtMsg_1 = new System.Windows.Forms.TextBox();
            this.VScroll1 = new System.Windows.Forms.VScrollBar();
            this.cmdAdd = new System.Windows.Forms.Button();
            this.txtMsg_10 = new System.Windows.Forms.TextBox();
            this.cmdX = new System.Windows.Forms.Button();
            this.cmdI = new System.Windows.Forms.Button();
            this.Line1 = new System.Windows.Forms.Label();
            this.lblMsg_10 = new System.Windows.Forms.Label();
            this.lblMsg_1 = new System.Windows.Forms.Label();
            this.lblMsg_2 = new System.Windows.Forms.Label();
            this.lblMsg_3 = new System.Windows.Forms.Label();
            this.lblMsg_4 = new System.Windows.Forms.Label();
            this.lblMsg_5 = new System.Windows.Forms.Label();
            this.lblMsg_6 = new System.Windows.Forms.Label();
            this.lblMsg_7 = new System.Windows.Forms.Label();
            this.lblMsg_8 = new System.Windows.Forms.Label();
            this.lblMsg_9 = new System.Windows.Forms.Label();
            this.SSTab1_Page2 = new System.Windows.Forms.TabPage();
            this.ListView1 = new System.Windows.Forms.ListView();
            this.ListView1_Col1 = new System.Windows.Forms.ColumnHeader();
            this.ListView1_Col2 = new System.Windows.Forms.ColumnHeader();
            this.SSTab1_Page3 = new System.Windows.Forms.TabPage();
            this.TreeView = new System.Windows.Forms.TreeView();
            this.Timer1 = new System.Windows.Forms.Timer(this.components);
            this.tmrVar = new System.Windows.Forms.Timer(this.components);
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.SSTab1.SuspendLayout();
            this.SSTab1_Page0.SuspendLayout();
            this.SSTab1_Page1.SuspendLayout();
            this.SSTab1_Page2.SuspendLayout();
            this.SSTab1_Page3.SuspendLayout();
            this.SuspendLayout();
            // 
            // Placeholder
            // 
            this.Placeholder.Location = new System.Drawing.Point(-67, 360);
            this.Placeholder.Name = "Placeholder";
            this.Placeholder.Size = new System.Drawing.Size(17, 13);
            this.Placeholder.TabIndex = 1;
            this.Placeholder.Text = "Check1";
            this.Placeholder.UseVisualStyleBackColor = true;
            // 
            // cmdCancel
            // 
            this.cmdCancel.Location = new System.Drawing.Point(296, 368);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(65, 25);
            this.cmdCancel.TabIndex = 3;
            this.cmdCancel.Text = "Cancel";
            this.cmdCancel.UseVisualStyleBackColor = true;
            // 
            // cmdOK
            // 
            this.cmdOK.Location = new System.Drawing.Point(368, 368);
            this.cmdOK.Name = "cmdOK";
            this.cmdOK.Size = new System.Drawing.Size(65, 25);
            this.cmdOK.TabIndex = 2;
            this.cmdOK.Text = "OK";
            this.cmdOK.UseVisualStyleBackColor = true;
            // 
            // SSTab1
            // 
            this.SSTab1.Controls.Add(this.SSTab1_Page0);
            this.SSTab1.Controls.Add(this.SSTab1_Page1);
            this.SSTab1.Controls.Add(this.SSTab1_Page2);
            this.SSTab1.Controls.Add(this.SSTab1_Page3);
            this.SSTab1.Location = new System.Drawing.Point(0, 8);
            this.SSTab1.Name = "SSTab1";
            this.SSTab1.Size = new System.Drawing.Size(441, 353);
            this.SSTab1.TabIndex = 4;
            this.SSTab1.SelectedIndex = 3;
            // 
            // SSTab1_Page0
            // 
            this.SSTab1_Page0.Controls.Add(this.Check1);
            this.SSTab1_Page0.Controls.Add(this.txtScript);
            this.SSTab1_Page0.Controls.Add(this.lblLineNum);
            this.SSTab1_Page0.Location = new System.Drawing.Point(4, 22);
            this.SSTab1_Page0.Name = "SSTab1_Page0";
            this.SSTab1_Page0.Size = new System.Drawing.Size(433, 327);
            this.SSTab1_Page0.TabIndex = 0;
            this.SSTab1_Page0.Text = "Script";
            this.SSTab1_Page0.UseVisualStyleBackColor = true;
            // 
            // Check1
            // 
            this.Check1.Location = new System.Drawing.Point(316, 300);
            this.Check1.Name = "Check1";
            this.Check1.Size = new System.Drawing.Size(113, 13);
            this.Check1.TabIndex = 6;
            this.Check1.Checked = true;
            this.Check1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.Check1.Text = "Show Line Number";
            this.Check1.UseVisualStyleBackColor = true;
            // 
            // txtScript
            // 
            this.txtScript.Font = new System.Drawing.Font("Courier New", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtScript.Location = new System.Drawing.Point(4, 0);
            this.txtScript.Name = "txtScript";
            this.txtScript.Size = new System.Drawing.Size(425, 305);
            this.txtScript.TabIndex = 0;
            this.txtScript.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedBoth;
            this.txtScript.Text = "";
            // 
            // lblLineNum
            // 
            this.lblLineNum.Location = new System.Drawing.Point(4, 300);
            this.lblLineNum.Name = "lblLineNum";
            this.lblLineNum.Size = new System.Drawing.Size(153, 17);
            this.lblLineNum.TabIndex = 5;
            this.lblLineNum.BackColor = System.Drawing.Color.Transparent;
            this.lblLineNum.Text = "Line Number: ";
            // 
            // SSTab1_Page1
            // 
            this.SSTab1_Page1.Controls.Add(this.txtMsg_9);
            this.SSTab1_Page1.Controls.Add(this.txtMsg_8);
            this.SSTab1_Page1.Controls.Add(this.txtMsg_7);
            this.SSTab1_Page1.Controls.Add(this.txtMsg_6);
            this.SSTab1_Page1.Controls.Add(this.txtMsg_5);
            this.SSTab1_Page1.Controls.Add(this.txtMsg_4);
            this.SSTab1_Page1.Controls.Add(this.txtMsg_3);
            this.SSTab1_Page1.Controls.Add(this.txtMsg_2);
            this.SSTab1_Page1.Controls.Add(this.txtMsg_1);
            this.SSTab1_Page1.Controls.Add(this.VScroll1);
            this.SSTab1_Page1.Controls.Add(this.cmdAdd);
            this.SSTab1_Page1.Controls.Add(this.txtMsg_10);
            this.SSTab1_Page1.Controls.Add(this.cmdX);
            this.SSTab1_Page1.Controls.Add(this.cmdI);
            this.SSTab1_Page1.Controls.Add(this.Line1);
            this.SSTab1_Page1.Controls.Add(this.lblMsg_10);
            this.SSTab1_Page1.Controls.Add(this.lblMsg_1);
            this.SSTab1_Page1.Controls.Add(this.lblMsg_2);
            this.SSTab1_Page1.Controls.Add(this.lblMsg_3);
            this.SSTab1_Page1.Controls.Add(this.lblMsg_4);
            this.SSTab1_Page1.Controls.Add(this.lblMsg_5);
            this.SSTab1_Page1.Controls.Add(this.lblMsg_6);
            this.SSTab1_Page1.Controls.Add(this.lblMsg_7);
            this.SSTab1_Page1.Controls.Add(this.lblMsg_8);
            this.SSTab1_Page1.Controls.Add(this.lblMsg_9);
            this.SSTab1_Page1.Location = new System.Drawing.Point(4, 22);
            this.SSTab1_Page1.Name = "SSTab1_Page1";
            this.SSTab1_Page1.Size = new System.Drawing.Size(433, 327);
            this.SSTab1_Page1.TabIndex = 1;
            this.SSTab1_Page1.Text = "Messages";
            this.SSTab1_Page1.UseVisualStyleBackColor = true;
            // 
            // txtMsg_9
            // 
            this.txtMsg_9.Location = new System.Drawing.Point(52, 200);
            this.txtMsg_9.Name = "txtMsg_9";
            this.txtMsg_9.Size = new System.Drawing.Size(313, 19);
            this.txtMsg_9.TabIndex = 20;
            this.txtMsg_9.Visible = false;
            this.txtMsg_9.Text = "";
            // 
            // txtMsg_8
            // 
            this.txtMsg_8.Location = new System.Drawing.Point(52, 176);
            this.txtMsg_8.Name = "txtMsg_8";
            this.txtMsg_8.Size = new System.Drawing.Size(313, 19);
            this.txtMsg_8.TabIndex = 19;
            this.txtMsg_8.Visible = false;
            this.txtMsg_8.Text = "";
            // 
            // txtMsg_7
            // 
            this.txtMsg_7.Location = new System.Drawing.Point(52, 152);
            this.txtMsg_7.Name = "txtMsg_7";
            this.txtMsg_7.Size = new System.Drawing.Size(313, 19);
            this.txtMsg_7.TabIndex = 18;
            this.txtMsg_7.Visible = false;
            this.txtMsg_7.Text = "";
            // 
            // txtMsg_6
            // 
            this.txtMsg_6.Location = new System.Drawing.Point(52, 128);
            this.txtMsg_6.Name = "txtMsg_6";
            this.txtMsg_6.Size = new System.Drawing.Size(313, 19);
            this.txtMsg_6.TabIndex = 17;
            this.txtMsg_6.Visible = false;
            this.txtMsg_6.Text = "";
            // 
            // txtMsg_5
            // 
            this.txtMsg_5.Location = new System.Drawing.Point(52, 104);
            this.txtMsg_5.Name = "txtMsg_5";
            this.txtMsg_5.Size = new System.Drawing.Size(313, 19);
            this.txtMsg_5.TabIndex = 16;
            this.txtMsg_5.Visible = false;
            this.txtMsg_5.Text = "";
            // 
            // txtMsg_4
            // 
            this.txtMsg_4.Location = new System.Drawing.Point(52, 80);
            this.txtMsg_4.Name = "txtMsg_4";
            this.txtMsg_4.Size = new System.Drawing.Size(313, 19);
            this.txtMsg_4.TabIndex = 15;
            this.txtMsg_4.Visible = false;
            this.txtMsg_4.Text = "";
            // 
            // txtMsg_3
            // 
            this.txtMsg_3.Location = new System.Drawing.Point(52, 56);
            this.txtMsg_3.Name = "txtMsg_3";
            this.txtMsg_3.Size = new System.Drawing.Size(313, 19);
            this.txtMsg_3.TabIndex = 14;
            this.txtMsg_3.Visible = false;
            this.txtMsg_3.Text = "";
            // 
            // txtMsg_2
            // 
            this.txtMsg_2.Location = new System.Drawing.Point(52, 32);
            this.txtMsg_2.Name = "txtMsg_2";
            this.txtMsg_2.Size = new System.Drawing.Size(313, 19);
            this.txtMsg_2.TabIndex = 13;
            this.txtMsg_2.Visible = false;
            this.txtMsg_2.Text = "";
            // 
            // txtMsg_1
            // 
            this.txtMsg_1.Location = new System.Drawing.Point(52, 8);
            this.txtMsg_1.Name = "txtMsg_1";
            this.txtMsg_1.Size = new System.Drawing.Size(313, 19);
            this.txtMsg_1.TabIndex = 12;
            this.txtMsg_1.Visible = false;
            this.txtMsg_1.Text = "";
            // 
            // VScroll1
            // 
            this.VScroll1.Location = new System.Drawing.Point(412, 0);
            this.VScroll1.Name = "VScroll1";
            this.VScroll1.Size = new System.Drawing.Size(17, 249);
            this.VScroll1.TabIndex = 11;
            this.VScroll1.LargeChange = 10;
            this.VScroll1.Maximum = 9;
            this.VScroll1.Minimum = 0;
            this.VScroll1.SmallChange = 1;
            this.VScroll1.Value = 0;
            // 
            // cmdAdd
            // 
            this.cmdAdd.Location = new System.Drawing.Point(364, 264);
            this.cmdAdd.Name = "cmdAdd";
            this.cmdAdd.Size = new System.Drawing.Size(65, 25);
            this.cmdAdd.TabIndex = 10;
            this.cmdAdd.Text = "Add New";
            this.cmdAdd.UseVisualStyleBackColor = true;
            // 
            // txtMsg_10
            // 
            this.txtMsg_10.Location = new System.Drawing.Point(52, 223);
            this.txtMsg_10.Name = "txtMsg_10";
            this.txtMsg_10.Size = new System.Drawing.Size(313, 19);
            this.txtMsg_10.TabIndex = 9;
            this.txtMsg_10.Visible = false;
            this.txtMsg_10.Text = "";
            // 
            // cmdX
            // 
            this.cmdX.Location = new System.Drawing.Point(388, 8);
            this.cmdX.Name = "cmdX";
            this.cmdX.Size = new System.Drawing.Size(17, 17);
            this.cmdX.TabIndex = 8;
            this.toolTip.SetToolTip(this.cmdX, "Delete");
            this.cmdX.Text = "×";
            this.cmdX.UseVisualStyleBackColor = true;
            // 
            // cmdI
            // 
            this.cmdI.Location = new System.Drawing.Point(372, 8);
            this.cmdI.Name = "cmdI";
            this.cmdI.Size = new System.Drawing.Size(17, 17);
            this.cmdI.TabIndex = 7;
            this.toolTip.SetToolTip(this.cmdI, "Insert New");
            this.cmdI.Text = "+";
            this.cmdI.UseVisualStyleBackColor = true;
            // 
            // Line1
            // 
            this.Line1.Location = new System.Drawing.Point(0, 256);
            this.Line1.Name = "Line1";
            this.Line1.Size = new System.Drawing.Size(440, 1);
            this.Line1.BackColor = System.Drawing.SystemColors.WindowText;
            // 
            // lblMsg_10
            // 
            this.lblMsg_10.Location = new System.Drawing.Point(0, 224);
            this.lblMsg_10.Name = "lblMsg_10";
            this.lblMsg_10.Size = new System.Drawing.Size(73, 17);
            this.lblMsg_10.TabIndex = 30;
            this.lblMsg_10.Visible = false;
            this.lblMsg_10.BackColor = System.Drawing.Color.Transparent;
            this.lblMsg_10.Text = "0  - ";
            this.lblMsg_10.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lblMsg_1
            // 
            this.lblMsg_1.Location = new System.Drawing.Point(4, 8);
            this.lblMsg_1.Name = "lblMsg_1";
            this.lblMsg_1.Size = new System.Drawing.Size(41, 17);
            this.lblMsg_1.TabIndex = 29;
            this.lblMsg_1.Visible = false;
            this.lblMsg_1.BackColor = System.Drawing.Color.Transparent;
            this.lblMsg_1.Text = "0  - ";
            this.lblMsg_1.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lblMsg_2
            // 
            this.lblMsg_2.Location = new System.Drawing.Point(0, 32);
            this.lblMsg_2.Name = "lblMsg_2";
            this.lblMsg_2.Size = new System.Drawing.Size(73, 17);
            this.lblMsg_2.TabIndex = 28;
            this.lblMsg_2.Visible = false;
            this.lblMsg_2.BackColor = System.Drawing.Color.Transparent;
            this.lblMsg_2.Text = "0  - ";
            this.lblMsg_2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lblMsg_3
            // 
            this.lblMsg_3.Location = new System.Drawing.Point(0, 56);
            this.lblMsg_3.Name = "lblMsg_3";
            this.lblMsg_3.Size = new System.Drawing.Size(73, 17);
            this.lblMsg_3.TabIndex = 27;
            this.lblMsg_3.Visible = false;
            this.lblMsg_3.BackColor = System.Drawing.Color.Transparent;
            this.lblMsg_3.Text = "0  - ";
            this.lblMsg_3.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lblMsg_4
            // 
            this.lblMsg_4.Location = new System.Drawing.Point(0, 80);
            this.lblMsg_4.Name = "lblMsg_4";
            this.lblMsg_4.Size = new System.Drawing.Size(73, 17);
            this.lblMsg_4.TabIndex = 26;
            this.lblMsg_4.Visible = false;
            this.lblMsg_4.BackColor = System.Drawing.Color.Transparent;
            this.lblMsg_4.Text = "0  - ";
            this.lblMsg_4.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lblMsg_5
            // 
            this.lblMsg_5.Location = new System.Drawing.Point(0, 104);
            this.lblMsg_5.Name = "lblMsg_5";
            this.lblMsg_5.Size = new System.Drawing.Size(73, 17);
            this.lblMsg_5.TabIndex = 25;
            this.lblMsg_5.Visible = false;
            this.lblMsg_5.BackColor = System.Drawing.Color.Transparent;
            this.lblMsg_5.Text = "0  - ";
            this.lblMsg_5.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lblMsg_6
            // 
            this.lblMsg_6.Location = new System.Drawing.Point(0, 128);
            this.lblMsg_6.Name = "lblMsg_6";
            this.lblMsg_6.Size = new System.Drawing.Size(73, 17);
            this.lblMsg_6.TabIndex = 24;
            this.lblMsg_6.Visible = false;
            this.lblMsg_6.BackColor = System.Drawing.Color.Transparent;
            this.lblMsg_6.Text = "0  - ";
            this.lblMsg_6.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lblMsg_7
            // 
            this.lblMsg_7.Location = new System.Drawing.Point(0, 152);
            this.lblMsg_7.Name = "lblMsg_7";
            this.lblMsg_7.Size = new System.Drawing.Size(73, 17);
            this.lblMsg_7.TabIndex = 23;
            this.lblMsg_7.Visible = false;
            this.lblMsg_7.BackColor = System.Drawing.Color.Transparent;
            this.lblMsg_7.Text = "0  - ";
            this.lblMsg_7.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lblMsg_8
            // 
            this.lblMsg_8.Location = new System.Drawing.Point(0, 176);
            this.lblMsg_8.Name = "lblMsg_8";
            this.lblMsg_8.Size = new System.Drawing.Size(73, 17);
            this.lblMsg_8.TabIndex = 22;
            this.lblMsg_8.Visible = false;
            this.lblMsg_8.BackColor = System.Drawing.Color.Transparent;
            this.lblMsg_8.Text = "0  - ";
            this.lblMsg_8.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lblMsg_9
            // 
            this.lblMsg_9.Location = new System.Drawing.Point(0, 200);
            this.lblMsg_9.Name = "lblMsg_9";
            this.lblMsg_9.Size = new System.Drawing.Size(73, 17);
            this.lblMsg_9.TabIndex = 21;
            this.lblMsg_9.Visible = false;
            this.lblMsg_9.BackColor = System.Drawing.Color.Transparent;
            this.lblMsg_9.Text = "0  - ";
            this.lblMsg_9.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // SSTab1_Page2
            // 
            this.SSTab1_Page2.Controls.Add(this.ListView1);
            this.SSTab1_Page2.Location = new System.Drawing.Point(4, 22);
            this.SSTab1_Page2.Name = "SSTab1_Page2";
            this.SSTab1_Page2.Size = new System.Drawing.Size(433, 327);
            this.SSTab1_Page2.TabIndex = 2;
            this.SSTab1_Page2.Text = "Variables";
            this.SSTab1_Page2.UseVisualStyleBackColor = true;
            // 
            // ListView1
            // 
            this.ListView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
                        this.ListView1_Col1,
                        this.ListView1_Col2});
            this.ListView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.ListView1.FullRowSelect = true;
            this.ListView1.HideSelection = false;
            this.ListView1.MultiSelect = false;
            this.ListView1.Location = new System.Drawing.Point(4, 0);
            this.ListView1.Name = "ListView1";
            this.ListView1.Size = new System.Drawing.Size(425, 313);
            this.ListView1.TabIndex = 31;
            this.ListView1.UseCompatibleStateImageBehavior = false;
            this.ListView1.View = System.Windows.Forms.View.Details;
            // 
            // ListView1_Col1
            // 
            this.ListView1_Col1.Text = "Variable";
            this.ListView1_Col1.Width = 200;
            // 
            // ListView1_Col2
            // 
            this.ListView1_Col2.Text = "Value";
            this.ListView1_Col2.Width = 550;
            // 
            // SSTab1_Page3
            // 
            this.SSTab1_Page3.Controls.Add(this.TreeView);
            this.SSTab1_Page3.Location = new System.Drawing.Point(4, 22);
            this.SSTab1_Page3.Name = "SSTab1_Page3";
            this.SSTab1_Page3.Size = new System.Drawing.Size(433, 327);
            this.SSTab1_Page3.TabIndex = 3;
            this.SSTab1_Page3.Text = "Player Arrays";
            this.SSTab1_Page3.UseVisualStyleBackColor = true;
            // 
            // TreeView
            // 
            this.TreeView.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.TreeView.Location = new System.Drawing.Point(4, 0);
            this.TreeView.Name = "TreeView";
            this.TreeView.Size = new System.Drawing.Size(425, 313);
            this.TreeView.TabIndex = 32;
            // 
            // Timer1
            // 
            this.Timer1.Enabled = true;
            this.Timer1.Interval = 1;
            // 
            // tmrVar
            // 
            this.tmrVar.Enabled = true;
            this.tmrVar.Interval = 100;
            // 
            // ServerScript
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(441, 394);
            this.Controls.Add(this.Placeholder);
            this.Controls.Add(this.cmdCancel);
            this.Controls.Add(this.cmdOK);
            this.Controls.Add(this.SSTab1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ServerScript";
            this.Text = "Server Script";
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
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.CheckBox Placeholder;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.Button cmdOK;
        private System.Windows.Forms.TabControl SSTab1;
        private System.Windows.Forms.TabPage SSTab1_Page0;
        private System.Windows.Forms.CheckBox Check1;
        private System.Windows.Forms.RichTextBox txtScript;
        private System.Windows.Forms.Label lblLineNum;
        private System.Windows.Forms.TabPage SSTab1_Page1;
        private System.Windows.Forms.TextBox txtMsg_9;
        private System.Windows.Forms.TextBox txtMsg_8;
        private System.Windows.Forms.TextBox txtMsg_7;
        private System.Windows.Forms.TextBox txtMsg_6;
        private System.Windows.Forms.TextBox txtMsg_5;
        private System.Windows.Forms.TextBox txtMsg_4;
        private System.Windows.Forms.TextBox txtMsg_3;
        private System.Windows.Forms.TextBox txtMsg_2;
        private System.Windows.Forms.TextBox txtMsg_1;
        private System.Windows.Forms.VScrollBar VScroll1;
        private System.Windows.Forms.Button cmdAdd;
        private System.Windows.Forms.TextBox txtMsg_10;
        private System.Windows.Forms.Button cmdX;
        private System.Windows.Forms.Button cmdI;
        private System.Windows.Forms.Label Line1;
        private System.Windows.Forms.Label lblMsg_10;
        private System.Windows.Forms.Label lblMsg_1;
        private System.Windows.Forms.Label lblMsg_2;
        private System.Windows.Forms.Label lblMsg_3;
        private System.Windows.Forms.Label lblMsg_4;
        private System.Windows.Forms.Label lblMsg_5;
        private System.Windows.Forms.Label lblMsg_6;
        private System.Windows.Forms.Label lblMsg_7;
        private System.Windows.Forms.Label lblMsg_8;
        private System.Windows.Forms.Label lblMsg_9;
        private System.Windows.Forms.TabPage SSTab1_Page2;
        private System.Windows.Forms.ListView ListView1;
        private System.Windows.Forms.ColumnHeader ListView1_Col1;
        private System.Windows.Forms.ColumnHeader ListView1_Col2;
        private System.Windows.Forms.TabPage SSTab1_Page3;
        private System.Windows.Forms.TreeView TreeView;
        private System.Windows.Forms.Timer Timer1;
        private System.Windows.Forms.Timer tmrVar;
        private System.Windows.Forms.ToolTip toolTip;
    }
}
