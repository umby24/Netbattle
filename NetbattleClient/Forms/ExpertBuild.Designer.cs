namespace Netbattle.Forms
{
    partial class ExpertBuild
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
            this.HPSelect = new System.Windows.Forms.ComboBox();
            this.Lev = new System.Windows.Forms.TextBox();
            this.Command1 = new System.Windows.Forms.Button();
            this.Shiny = new System.Windows.Forms.CheckBox();
            this.DV_SAtk = new System.Windows.Forms.ComboBox();
            this.DV_Spd = new System.Windows.Forms.ComboBox();
            this.DV_Def = new System.Windows.Forms.ComboBox();
            this.DV_Atk = new System.Windows.Forms.ComboBox();
            this.Label3 = new System.Windows.Forms.Label();
            this.GenderDisp = new System.Windows.Forms.Label();
            this.HP = new System.Windows.Forms.Label();
            this.HiddenPower = new System.Windows.Forms.Label();
            this.SDefensePower = new System.Windows.Forms.Label();
            this.SAttackPower = new System.Windows.Forms.Label();
            this.Label6 = new System.Windows.Forms.Label();
            this.SpeedPower = new System.Windows.Forms.Label();
            this.Label4 = new System.Windows.Forms.Label();
            this.DefensePower = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.AttackPower = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // HPSelect
            // 
            this.HPSelect.Location = new System.Drawing.Point(24, 128);
            this.HPSelect.Name = "HPSelect";
            this.HPSelect.Size = new System.Drawing.Size(113, 21);
            this.HPSelect.TabIndex = 20;
            this.HPSelect.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.HPSelect.FormattingEnabled = true;
            // 
            // Lev
            // 
            this.Lev.Location = new System.Drawing.Point(200, 24);
            this.Lev.Name = "Lev";
            this.Lev.Size = new System.Drawing.Size(33, 19);
            this.Lev.TabIndex = 18;
            this.Lev.MaxLength = 3;
            this.Lev.Text = "100";
            // 
            // Command1
            // 
            this.Command1.Location = new System.Drawing.Point(160, 128);
            this.Command1.Name = "Command1";
            this.Command1.Size = new System.Drawing.Size(73, 25);
            this.Command1.TabIndex = 15;
            this.Command1.Text = "&Ok";
            this.Command1.UseVisualStyleBackColor = true;
            // 
            // Shiny
            // 
            this.Shiny.Location = new System.Drawing.Point(8, 80);
            this.Shiny.Name = "Shiny";
            this.Shiny.Size = new System.Drawing.Size(49, 17);
            this.Shiny.TabIndex = 13;
            this.Shiny.Enabled = false;
            this.Shiny.Text = "Shiny";
            this.Shiny.UseVisualStyleBackColor = true;
            // 
            // DV_SAtk
            // 
            this.DV_SAtk.Location = new System.Drawing.Point(152, 24);
            this.DV_SAtk.Name = "DV_SAtk";
            this.DV_SAtk.Size = new System.Drawing.Size(41, 21);
            this.DV_SAtk.TabIndex = 9;
            this.DV_SAtk.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.DV_SAtk.FormattingEnabled = true;
            this.DV_SAtk.Items.AddRange(new object[] {
                        "0",
                        "1",
                        "2",
                        "3",
                        "4",
                        "5",
                        "6",
                        "7",
                        "8",
                        "9",
                        "10",
                        "11",
                        "12",
                        "13",
                        "14",
                        "15"});
            // 
            // DV_Spd
            // 
            this.DV_Spd.Location = new System.Drawing.Point(104, 24);
            this.DV_Spd.Name = "DV_Spd";
            this.DV_Spd.Size = new System.Drawing.Size(41, 21);
            this.DV_Spd.TabIndex = 6;
            this.DV_Spd.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.DV_Spd.FormattingEnabled = true;
            this.DV_Spd.Items.AddRange(new object[] {
                        "0",
                        "1",
                        "2",
                        "3",
                        "4",
                        "5",
                        "6",
                        "7",
                        "8",
                        "9",
                        "10",
                        "11",
                        "12",
                        "13",
                        "14",
                        "15"});
            // 
            // DV_Def
            // 
            this.DV_Def.Location = new System.Drawing.Point(56, 24);
            this.DV_Def.Name = "DV_Def";
            this.DV_Def.Size = new System.Drawing.Size(41, 21);
            this.DV_Def.TabIndex = 3;
            this.DV_Def.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.DV_Def.FormattingEnabled = true;
            this.DV_Def.Items.AddRange(new object[] {
                        "0",
                        "1",
                        "2",
                        "3",
                        "4",
                        "5",
                        "6",
                        "7",
                        "8",
                        "9",
                        "10",
                        "11",
                        "12",
                        "13",
                        "14",
                        "15"});
            // 
            // DV_Atk
            // 
            this.DV_Atk.Location = new System.Drawing.Point(8, 24);
            this.DV_Atk.Name = "DV_Atk";
            this.DV_Atk.Size = new System.Drawing.Size(41, 21);
            this.DV_Atk.TabIndex = 0;
            this.DV_Atk.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.DV_Atk.FormattingEnabled = true;
            this.DV_Atk.Items.AddRange(new object[] {
                        "\u0001",
                        "\u0001",
                        "\u0001",
                        "\u0001",
                        "\u0001",
                        "\u0001",
                        "\u0001",
                        "\u0001",
                        "\u0001",
                        "\u0001",
                        "10",
                        "11",
                        "12",
                        "13",
                        "14",
                        "15"});
            // 
            // Label3
            // 
            this.Label3.Location = new System.Drawing.Point(200, 8);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(41, 17);
            this.Label3.TabIndex = 19;
            this.Label3.BackColor = System.Drawing.Color.Transparent;
            this.Label3.Text = "Level";
            // 
            // GenderDisp
            // 
            this.GenderDisp.Location = new System.Drawing.Point(128, 96);
            this.GenderDisp.Name = "GenderDisp";
            this.GenderDisp.Size = new System.Drawing.Size(105, 17);
            this.GenderDisp.TabIndex = 17;
            this.GenderDisp.BackColor = System.Drawing.Color.Transparent;
            this.GenderDisp.Text = "Gender:";
            // 
            // HP
            // 
            this.HP.Location = new System.Drawing.Point(24, 96);
            this.HP.Name = "HP";
            this.HP.Size = new System.Drawing.Size(65, 17);
            this.HP.TabIndex = 16;
            this.HP.BackColor = System.Drawing.Color.Transparent;
            this.HP.Text = "HP:";
            // 
            // HiddenPower
            // 
            this.HiddenPower.Location = new System.Drawing.Point(24, 112);
            this.HiddenPower.Name = "HiddenPower";
            this.HiddenPower.Size = new System.Drawing.Size(169, 17);
            this.HiddenPower.TabIndex = 14;
            this.HiddenPower.BackColor = System.Drawing.Color.Transparent;
            this.HiddenPower.Text = "Hidden Power:";
            // 
            // SDefensePower
            // 
            this.SDefensePower.Location = new System.Drawing.Point(152, 64);
            this.SDefensePower.Name = "SDefensePower";
            this.SDefensePower.Size = new System.Drawing.Size(41, 17);
            this.SDefensePower.TabIndex = 12;
            this.SDefensePower.BackColor = System.Drawing.Color.Transparent;
            this.SDefensePower.Text = "";
            // 
            // SAttackPower
            // 
            this.SAttackPower.Location = new System.Drawing.Point(152, 48);
            this.SAttackPower.Name = "SAttackPower";
            this.SAttackPower.Size = new System.Drawing.Size(41, 17);
            this.SAttackPower.TabIndex = 11;
            this.SAttackPower.BackColor = System.Drawing.Color.Transparent;
            this.SAttackPower.Text = "";
            // 
            // Label6
            // 
            this.Label6.Location = new System.Drawing.Point(152, 8);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(41, 17);
            this.Label6.TabIndex = 10;
            this.Label6.BackColor = System.Drawing.Color.Transparent;
            this.Label6.Text = "Special";
            // 
            // SpeedPower
            // 
            this.SpeedPower.Location = new System.Drawing.Point(104, 48);
            this.SpeedPower.Name = "SpeedPower";
            this.SpeedPower.Size = new System.Drawing.Size(41, 17);
            this.SpeedPower.TabIndex = 8;
            this.SpeedPower.BackColor = System.Drawing.Color.Transparent;
            this.SpeedPower.Text = "";
            // 
            // Label4
            // 
            this.Label4.Location = new System.Drawing.Point(104, 8);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(41, 17);
            this.Label4.TabIndex = 7;
            this.Label4.BackColor = System.Drawing.Color.Transparent;
            this.Label4.Text = "Speed";
            // 
            // DefensePower
            // 
            this.DefensePower.Location = new System.Drawing.Point(56, 48);
            this.DefensePower.Name = "DefensePower";
            this.DefensePower.Size = new System.Drawing.Size(41, 17);
            this.DefensePower.TabIndex = 5;
            this.DefensePower.BackColor = System.Drawing.Color.Transparent;
            this.DefensePower.Text = "";
            // 
            // Label2
            // 
            this.Label2.Location = new System.Drawing.Point(56, 8);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(41, 17);
            this.Label2.TabIndex = 4;
            this.Label2.BackColor = System.Drawing.Color.Transparent;
            this.Label2.Text = "Defense";
            // 
            // AttackPower
            // 
            this.AttackPower.Location = new System.Drawing.Point(8, 48);
            this.AttackPower.Name = "AttackPower";
            this.AttackPower.Size = new System.Drawing.Size(41, 17);
            this.AttackPower.TabIndex = 2;
            this.AttackPower.BackColor = System.Drawing.Color.Transparent;
            this.AttackPower.Text = "";
            // 
            // Label1
            // 
            this.Label1.Location = new System.Drawing.Point(8, 8);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(41, 17);
            this.Label1.TabIndex = 1;
            this.Label1.BackColor = System.Drawing.Color.Transparent;
            this.Label1.Text = "Attack";
            // 
            // ExpertBuild
            // 
            this.AcceptButton = this.Command1;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.Command1;
            this.ClientSize = new System.Drawing.Size(242, 161);
            this.ControlBox = false;
            this.Controls.Add(this.HPSelect);
            this.Controls.Add(this.Lev);
            this.Controls.Add(this.Command1);
            this.Controls.Add(this.Shiny);
            this.Controls.Add(this.DV_SAtk);
            this.Controls.Add(this.DV_Spd);
            this.Controls.Add(this.DV_Def);
            this.Controls.Add(this.DV_Atk);
            this.Controls.Add(this.Label3);
            this.Controls.Add(this.GenderDisp);
            this.Controls.Add(this.HP);
            this.Controls.Add(this.HiddenPower);
            this.Controls.Add(this.SDefensePower);
            this.Controls.Add(this.SAttackPower);
            this.Controls.Add(this.Label6);
            this.Controls.Add(this.SpeedPower);
            this.Controls.Add(this.Label4);
            this.Controls.Add(this.DefensePower);
            this.Controls.Add(this.Label2);
            this.Controls.Add(this.AttackPower);
            this.Controls.Add(this.Label1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ExpertBuild";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Expert Build";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.ComboBox HPSelect;
        private System.Windows.Forms.TextBox Lev;
        private System.Windows.Forms.Button Command1;
        private System.Windows.Forms.CheckBox Shiny;
        private System.Windows.Forms.ComboBox DV_SAtk;
        private System.Windows.Forms.ComboBox DV_Spd;
        private System.Windows.Forms.ComboBox DV_Def;
        private System.Windows.Forms.ComboBox DV_Atk;
        private System.Windows.Forms.Label Label3;
        private System.Windows.Forms.Label GenderDisp;
        private System.Windows.Forms.Label HP;
        private System.Windows.Forms.Label HiddenPower;
        private System.Windows.Forms.Label SDefensePower;
        private System.Windows.Forms.Label SAttackPower;
        private System.Windows.Forms.Label Label6;
        private System.Windows.Forms.Label SpeedPower;
        private System.Windows.Forms.Label Label4;
        private System.Windows.Forms.Label DefensePower;
        private System.Windows.Forms.Label Label2;
        private System.Windows.Forms.Label AttackPower;
        private System.Windows.Forms.Label Label1;
    }
}
