namespace Netbattle.Forms
{
    partial class PokedexSearch
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PokedexSearch));
            this.Frame2 = new System.Windows.Forms.GroupBox();
            this.Results = new System.Windows.Forms.ListView();
            this.Results_Col1 = new System.Windows.Forms.ColumnHeader();
            this.Results_Col2 = new System.Windows.Forms.ColumnHeader();
            this.Results_Col3 = new System.Windows.Forms.ColumnHeader();
            this.Results_Col4 = new System.Windows.Forms.ColumnHeader();
            this.Results_Col5 = new System.Windows.Forms.ColumnHeader();
            this.Results_Col6 = new System.Windows.Forms.ColumnHeader();
            this.Results_Col7 = new System.Windows.Forms.ColumnHeader();
            this.Results_Col8 = new System.Windows.Forms.ColumnHeader();
            this.Results_Col9 = new System.Windows.Forms.ColumnHeader();
            this.Frame1 = new System.Windows.Forms.GroupBox();
            this.Picture1 = new System.Windows.Forms.Panel();
            this.MoveList = new System.Windows.Forms.ListView();
            this.MoveList_Col1 = new System.Windows.Forms.ColumnHeader();
            this.cmbSearch_3 = new System.Windows.Forms.ComboBox();
            this.NameBox = new System.Windows.Forms.TextBox();
            this.cmbSearch_1 = new System.Windows.Forms.ComboBox();
            this.cmbSearch_4 = new System.Windows.Forms.ComboBox();
            this.cmbSearch_5 = new System.Windows.Forms.ComboBox();
            this.cmbSearch_2 = new System.Windows.Forms.ComboBox();
            this.SCheck_0 = new System.Windows.Forms.CheckBox();
            this.CancelButton_ = new System.Windows.Forms.Button();
            this.SearchButton = new System.Windows.Forms.Button();
            this.SCheck_9 = new System.Windows.Forms.CheckBox();
            this.SCheck_2 = new System.Windows.Forms.CheckBox();
            this.SCheck_1 = new System.Windows.Forms.CheckBox();
            this.SCheck_3 = new System.Windows.Forms.CheckBox();
            this.SCheck_4 = new System.Windows.Forms.CheckBox();
            this.SCheck_5 = new System.Windows.Forms.CheckBox();
            this.Frame2.SuspendLayout();
            this.Frame1.SuspendLayout();
            this.Picture1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Frame2
            // 
            this.Frame2.Controls.Add(this.Results);
            this.Frame2.Location = new System.Drawing.Point(8, 200);
            this.Frame2.Name = "Frame2";
            this.Frame2.Size = new System.Drawing.Size(409, 209);
            this.Frame2.TabIndex = 16;
            this.Frame2.TabStop = false;
            this.Frame2.Text = "Search Results";
            // 
            // Results
            // 
            this.Results.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
                        this.Results_Col1,
                        this.Results_Col2,
                        this.Results_Col3,
                        this.Results_Col4,
                        this.Results_Col5,
                        this.Results_Col6,
                        this.Results_Col7,
                        this.Results_Col8,
                        this.Results_Col9});
            this.Results.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.Results.FullRowSelect = true;
            this.Results.MultiSelect = false;
            this.Results.Location = new System.Drawing.Point(8, 16);
            this.Results.Name = "Results";
            this.Results.Size = new System.Drawing.Size(393, 185);
            this.Results.TabIndex = 17;
            this.Results.UseCompatibleStateImageBehavior = false;
            this.Results.View = System.Windows.Forms.View.Details;
            // 
            // Results_Col1
            // 
            this.Results_Col1.Text = "Name";
            this.Results_Col1.Width = 122;
            // 
            // Results_Col2
            // 
            this.Results_Col2.Text = "Type1";
            this.Results_Col2.Width = 86;
            // 
            // Results_Col3
            // 
            this.Results_Col3.Text = "Type2";
            this.Results_Col3.Width = 86;
            // 
            // Results_Col4
            // 
            this.Results_Col4.Text = "HP";
            this.Results_Col4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.Results_Col4.Width = 58;
            // 
            // Results_Col5
            // 
            this.Results_Col5.Text = "Atk";
            this.Results_Col5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.Results_Col5.Width = 58;
            // 
            // Results_Col6
            // 
            this.Results_Col6.Text = "Def";
            this.Results_Col6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.Results_Col6.Width = 58;
            // 
            // Results_Col7
            // 
            this.Results_Col7.Text = "Spd";
            this.Results_Col7.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.Results_Col7.Width = 58;
            // 
            // Results_Col8
            // 
            this.Results_Col8.Text = "SAtk";
            this.Results_Col8.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.Results_Col8.Width = 64;
            // 
            // Results_Col9
            // 
            this.Results_Col9.Text = "SDef";
            this.Results_Col9.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.Results_Col9.Width = 64;
            // 
            // Frame1
            // 
            this.Frame1.Controls.Add(this.Picture1);
            this.Frame1.Location = new System.Drawing.Point(8, 0);
            this.Frame1.Name = "Frame1";
            this.Frame1.Size = new System.Drawing.Size(409, 193);
            this.Frame1.TabIndex = 0;
            this.Frame1.TabStop = false;
            this.Frame1.Text = "Search Options";
            // 
            // Picture1
            // 
            this.Picture1.Controls.Add(this.MoveList);
            this.Picture1.Controls.Add(this.cmbSearch_3);
            this.Picture1.Controls.Add(this.NameBox);
            this.Picture1.Controls.Add(this.cmbSearch_1);
            this.Picture1.Controls.Add(this.cmbSearch_4);
            this.Picture1.Controls.Add(this.cmbSearch_5);
            this.Picture1.Controls.Add(this.cmbSearch_2);
            this.Picture1.Controls.Add(this.SCheck_0);
            this.Picture1.Controls.Add(this.CancelButton_);
            this.Picture1.Controls.Add(this.SearchButton);
            this.Picture1.Controls.Add(this.SCheck_9);
            this.Picture1.Controls.Add(this.SCheck_2);
            this.Picture1.Controls.Add(this.SCheck_1);
            this.Picture1.Controls.Add(this.SCheck_3);
            this.Picture1.Controls.Add(this.SCheck_4);
            this.Picture1.Controls.Add(this.SCheck_5);
            this.Picture1.Location = new System.Drawing.Point(8, 16);
            this.Picture1.Name = "Picture1";
            this.Picture1.Size = new System.Drawing.Size(393, 169);
            this.Picture1.TabIndex = 1;
            // 
            // MoveList
            // 
            this.MoveList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
                        this.MoveList_Col1});
            this.MoveList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.MoveList.CheckBoxes = true;
            this.MoveList.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.MoveList.HideSelection = false;
            this.MoveList.MultiSelect = false;
            this.MoveList.Location = new System.Drawing.Point(240, 16);
            this.MoveList.Name = "MoveList";
            this.MoveList.Size = new System.Drawing.Size(153, 121);
            this.MoveList.TabIndex = 15;
            this.MoveList.UseCompatibleStateImageBehavior = false;
            this.MoveList.View = System.Windows.Forms.View.Details;
            // 
            // MoveList_Col1
            // 
            this.MoveList_Col1.Text = "Move";
            this.MoveList_Col1.Width = 339;
            // 
            // cmbSearch_3
            // 
            this.cmbSearch_3.Location = new System.Drawing.Point(120, 16);
            this.cmbSearch_3.Name = "cmbSearch_3";
            this.cmbSearch_3.Size = new System.Drawing.Size(113, 21);
            this.cmbSearch_3.TabIndex = 9;
            this.cmbSearch_3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSearch_3.FormattingEnabled = true;
            this.cmbSearch_3.Sorted = true;
            // 
            // NameBox
            // 
            this.NameBox.Location = new System.Drawing.Point(0, 16);
            this.NameBox.Name = "NameBox";
            this.NameBox.Size = new System.Drawing.Size(113, 21);
            this.NameBox.TabIndex = 19;
            this.NameBox.Text = "Text1";
            // 
            // cmbSearch_1
            // 
            this.cmbSearch_1.Location = new System.Drawing.Point(0, 64);
            this.cmbSearch_1.Name = "cmbSearch_1";
            this.cmbSearch_1.Size = new System.Drawing.Size(113, 21);
            this.cmbSearch_1.TabIndex = 7;
            this.cmbSearch_1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSearch_1.FormattingEnabled = true;
            // 
            // cmbSearch_4
            // 
            this.cmbSearch_4.Location = new System.Drawing.Point(120, 64);
            this.cmbSearch_4.Name = "cmbSearch_4";
            this.cmbSearch_4.Size = new System.Drawing.Size(113, 21);
            this.cmbSearch_4.TabIndex = 5;
            this.cmbSearch_4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSearch_4.FormattingEnabled = true;
            // 
            // cmbSearch_5
            // 
            this.cmbSearch_5.Location = new System.Drawing.Point(120, 112);
            this.cmbSearch_5.Name = "cmbSearch_5";
            this.cmbSearch_5.Size = new System.Drawing.Size(113, 21);
            this.cmbSearch_5.TabIndex = 4;
            this.cmbSearch_5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSearch_5.FormattingEnabled = true;
            // 
            // cmbSearch_2
            // 
            this.cmbSearch_2.Location = new System.Drawing.Point(0, 112);
            this.cmbSearch_2.Name = "cmbSearch_2";
            this.cmbSearch_2.Size = new System.Drawing.Size(113, 21);
            this.cmbSearch_2.TabIndex = 6;
            this.cmbSearch_2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSearch_2.FormattingEnabled = true;
            // 
            // SCheck_0
            // 
            this.SCheck_0.Location = new System.Drawing.Point(0, 0);
            this.SCheck_0.Name = "SCheck_0";
            this.SCheck_0.Size = new System.Drawing.Size(65, 17);
            this.SCheck_0.TabIndex = 18;
            this.SCheck_0.Text = "Name";
            this.SCheck_0.UseVisualStyleBackColor = true;
            // 
            // CancelButton_
            // 
            this.CancelButton_.Location = new System.Drawing.Point(320, 144);
            this.CancelButton_.Name = "CancelButton_";
            this.CancelButton_.Size = new System.Drawing.Size(73, 25);
            this.CancelButton_.TabIndex = 14;
            this.CancelButton_.Text = "&Exit";
            this.CancelButton_.UseVisualStyleBackColor = true;
            // 
            // SearchButton
            // 
            this.SearchButton.Location = new System.Drawing.Point(240, 144);
            this.SearchButton.Name = "SearchButton";
            this.SearchButton.Size = new System.Drawing.Size(73, 25);
            this.SearchButton.TabIndex = 13;
            this.SearchButton.Enabled = false;
            this.SearchButton.Text = "&Search";
            this.SearchButton.UseVisualStyleBackColor = true;
            // 
            // SCheck_9
            // 
            this.SCheck_9.Location = new System.Drawing.Point(240, 0);
            this.SCheck_9.Name = "SCheck_9";
            this.SCheck_9.Size = new System.Drawing.Size(65, 17);
            this.SCheck_9.TabIndex = 12;
            this.SCheck_9.Text = "Moves";
            this.SCheck_9.UseVisualStyleBackColor = true;
            // 
            // SCheck_2
            // 
            this.SCheck_2.Location = new System.Drawing.Point(0, 96);
            this.SCheck_2.Name = "SCheck_2";
            this.SCheck_2.Size = new System.Drawing.Size(65, 17);
            this.SCheck_2.TabIndex = 11;
            this.SCheck_2.Text = "2nd Type";
            this.SCheck_2.UseVisualStyleBackColor = true;
            // 
            // SCheck_1
            // 
            this.SCheck_1.Location = new System.Drawing.Point(0, 48);
            this.SCheck_1.Name = "SCheck_1";
            this.SCheck_1.Size = new System.Drawing.Size(65, 17);
            this.SCheck_1.TabIndex = 10;
            this.SCheck_1.Text = "Type";
            this.SCheck_1.UseVisualStyleBackColor = true;
            // 
            // SCheck_3
            // 
            this.SCheck_3.Location = new System.Drawing.Point(120, 0);
            this.SCheck_3.Name = "SCheck_3";
            this.SCheck_3.Size = new System.Drawing.Size(65, 17);
            this.SCheck_3.TabIndex = 8;
            this.SCheck_3.Text = "Trait";
            this.SCheck_3.UseVisualStyleBackColor = true;
            // 
            // SCheck_4
            // 
            this.SCheck_4.Location = new System.Drawing.Point(120, 48);
            this.SCheck_4.Name = "SCheck_4";
            this.SCheck_4.Size = new System.Drawing.Size(81, 17);
            this.SCheck_4.TabIndex = 3;
            this.SCheck_4.Text = "Egg Group";
            this.SCheck_4.UseVisualStyleBackColor = true;
            // 
            // SCheck_5
            // 
            this.SCheck_5.Location = new System.Drawing.Point(120, 96);
            this.SCheck_5.Name = "SCheck_5";
            this.SCheck_5.Size = new System.Drawing.Size(105, 17);
            this.SCheck_5.TabIndex = 2;
            this.SCheck_5.Text = "2nd Egg Group";
            this.SCheck_5.UseVisualStyleBackColor = true;
            // 
            // PokedexSearch
            // 
            this.AcceptButton = this.SearchButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CancelButton_;
            this.ClientSize = new System.Drawing.Size(421, 416);
            this.Controls.Add(this.Frame2);
            this.Controls.Add(this.Frame1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PokedexSearch";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Search Pokedex";
            this.Frame2.ResumeLayout(false);
            this.Frame2.PerformLayout();
            this.Frame1.ResumeLayout(false);
            this.Frame1.PerformLayout();
            this.Picture1.ResumeLayout(false);
            this.Picture1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.GroupBox Frame2;
        private System.Windows.Forms.ListView Results;
        private System.Windows.Forms.ColumnHeader Results_Col1;
        private System.Windows.Forms.ColumnHeader Results_Col2;
        private System.Windows.Forms.ColumnHeader Results_Col3;
        private System.Windows.Forms.ColumnHeader Results_Col4;
        private System.Windows.Forms.ColumnHeader Results_Col5;
        private System.Windows.Forms.ColumnHeader Results_Col6;
        private System.Windows.Forms.ColumnHeader Results_Col7;
        private System.Windows.Forms.ColumnHeader Results_Col8;
        private System.Windows.Forms.ColumnHeader Results_Col9;
        private System.Windows.Forms.GroupBox Frame1;
        private System.Windows.Forms.Panel Picture1;
        private System.Windows.Forms.ListView MoveList;
        private System.Windows.Forms.ColumnHeader MoveList_Col1;
        private System.Windows.Forms.ComboBox cmbSearch_3;
        private System.Windows.Forms.TextBox NameBox;
        private System.Windows.Forms.ComboBox cmbSearch_1;
        private System.Windows.Forms.ComboBox cmbSearch_4;
        private System.Windows.Forms.ComboBox cmbSearch_5;
        private System.Windows.Forms.ComboBox cmbSearch_2;
        private System.Windows.Forms.CheckBox SCheck_0;
        private System.Windows.Forms.Button CancelButton_;
        private System.Windows.Forms.Button SearchButton;
        private System.Windows.Forms.CheckBox SCheck_9;
        private System.Windows.Forms.CheckBox SCheck_2;
        private System.Windows.Forms.CheckBox SCheck_1;
        private System.Windows.Forms.CheckBox SCheck_3;
        private System.Windows.Forms.CheckBox SCheck_4;
        private System.Windows.Forms.CheckBox SCheck_5;
    }
}
