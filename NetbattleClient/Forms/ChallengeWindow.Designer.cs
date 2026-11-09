namespace Netbattle.Forms {
    partial class ChallengeWindow {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
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
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChallengeWindow));
            lblWins = new Label();
            lblLosses = new Label();
            lblTies = new Label();
            lblUbers = new Label();
            ruleColumn = new ColumnHeader();
            toolTip1 = new ToolTip(components);
            tmrRandbat = new System.Windows.Forms.Timer(components);
            btnChallenge = new Button();
            btnCancel = new Button();
            listView1 = new ListView();
            lblClauses = new Label();
            grpTeam = new GroupBox();
            panel3 = new Panel();
            panel2 = new Panel();
            panel1 = new Panel();
            pictureBox1 = new PictureBox();
            label3 = new Label();
            label2 = new Label();
            label1 = new Label();
            lblEstPower = new Label();
            progressBar1 = new ProgressBar();
            lblExtraInfo = new Label();
            pk6 = new PictureBox();
            pk5 = new PictureBox();
            pk4 = new PictureBox();
            pk3 = new PictureBox();
            pk2 = new PictureBox();
            pk1 = new PictureBox();
            grpRulesModes = new GroupBox();
            lblTerrain = new Label();
            lblMode = new Label();
            comboBox2 = new ComboBox();
            comboBox1 = new ComboBox();
            label4 = new Label();
            grpTeam.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pk6).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pk5).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pk4).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pk3).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pk2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pk1).BeginInit();
            grpRulesModes.SuspendLayout();
            SuspendLayout();
            // 
            // btnChallenge
            // 
            btnChallenge.Location = new Point(13, 413);
            btnChallenge.Margin = new Padding(2);
            btnChallenge.Name = "btnChallenge";
            btnChallenge.Size = new Size(166, 22);
            btnChallenge.TabIndex = 0;
            btnChallenge.Text = "Challenge!";
            btnChallenge.UseVisualStyleBackColor = true;
            btnChallenge.Click += btnChallenge_Click;
            // 
            // btnCancel
            // 
            btnCancel.Location = new Point(183, 413);
            btnCancel.Margin = new Padding(2);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(158, 22);
            btnCancel.TabIndex = 1;
            btnCancel.Text = "Cancel";
            btnCancel.UseVisualStyleBackColor = true;
            btnCancel.Click += btnCancel_Click;
            // 
            // listView1
            // 
            listView1.Columns.AddRange(new ColumnHeader[] { ruleColumn });
            listView1.FullRowSelect = true;
            listView1.HeaderStyle = ColumnHeaderStyle.None;
            listView1.Location = new Point(6, 82);
            listView1.Margin = new Padding(2);
            listView1.MultiSelect = false;
            listView1.Name = "listView1";
            listView1.ShowItemToolTips = true;
            listView1.Size = new Size(324, 92);
            listView1.TabIndex = 2;
            listView1.UseCompatibleStateImageBehavior = false;
            listView1.View = View.Details;
            listView1.ItemCheck += listView1_ItemCheck;
            listView1.ItemChecked += listView1_ItemChecked;
            listView1.SelectedIndexChanged += listView1_SelectedIndexChanged;
            //
            // ruleColumn
            //
            ruleColumn.Width = 300;
            // 
            // lblClauses
            // 
            lblClauses.AutoSize = true;
            lblClauses.Location = new Point(6, 65);
            lblClauses.Margin = new Padding(2, 0, 2, 0);
            lblClauses.Name = "lblClauses";
            lblClauses.Size = new Size(47, 15);
            lblClauses.TabIndex = 3;
            lblClauses.Text = "Clauses";
            // 
            // grpTeam
            // 
            grpTeam.Controls.Add(lblUbers);
            grpTeam.Controls.Add(panel3);
            grpTeam.Controls.Add(panel2);
            grpTeam.Controls.Add(panel1);
            grpTeam.Controls.Add(pictureBox1);
            grpTeam.Controls.Add(label3);
            grpTeam.Controls.Add(label2);
            grpTeam.Controls.Add(label1);
            grpTeam.Controls.Add(lblEstPower);
            grpTeam.Controls.Add(progressBar1);
            grpTeam.Controls.Add(lblExtraInfo);
            grpTeam.Controls.Add(pk6);
            grpTeam.Controls.Add(pk5);
            grpTeam.Controls.Add(pk4);
            grpTeam.Controls.Add(pk3);
            grpTeam.Controls.Add(pk2);
            grpTeam.Controls.Add(pk1);
            grpTeam.Location = new Point(13, 45);
            grpTeam.Margin = new Padding(2);
            grpTeam.Name = "grpTeam";
            grpTeam.Padding = new Padding(2);
            grpTeam.Size = new Size(331, 168);
            grpTeam.TabIndex = 4;
            grpTeam.TabStop = false;
            grpTeam.Text = "Player and Team Info";
            // 
            // panel3
            // 
            panel3.BackColor = Color.FromArgb(255, 255, 192);
            panel3.BorderStyle = BorderStyle.Fixed3D;
            panel3.Controls.Add(lblTies);
            panel3.Location = new Point(274, 128);
            panel3.Margin = new Padding(4, 3, 4, 3);
            panel3.Name = "panel3";
            panel3.Size = new Size(38, 37);
            panel3.TabIndex = 15;
            // 
            // panel2
            // 
            panel2.BackColor = Color.FromArgb(255, 192, 192);
            panel2.BorderStyle = BorderStyle.Fixed3D;
            panel2.Controls.Add(lblLosses);
            panel2.Location = new Point(229, 128);
            panel2.Margin = new Padding(4, 3, 4, 3);
            panel2.Name = "panel2";
            panel2.Size = new Size(38, 37);
            panel2.TabIndex = 14;
            // 
            // panel1
            // 
            panel1.BackColor = Color.FromArgb(192, 255, 192);
            panel1.BorderStyle = BorderStyle.Fixed3D;
            panel1.Controls.Add(lblWins);
            panel1.Location = new Point(183, 128);
            panel1.Margin = new Padding(4, 3, 4, 3);
            panel1.Name = "panel1";
            panel1.Size = new Size(38, 37);
            panel1.TabIndex = 13;
            // 
            // pictureBox1
            // 
            pictureBox1.Location = new Point(5, 122);
            pictureBox1.Margin = new Padding(2);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(37, 37);
            pictureBox1.TabIndex = 12;
            pictureBox1.TabStop = false;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(278, 110);
            label3.Margin = new Padding(2, 0, 2, 0);
            label3.Name = "label3";
            label3.Size = new Size(30, 15);
            label3.TabIndex = 11;
            label3.Text = "Ties:";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(229, 110);
            label2.Margin = new Padding(2, 0, 2, 0);
            label2.Name = "label2";
            label2.Size = new Size(44, 15);
            label2.TabIndex = 10;
            label2.Text = "Losses:";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(182, 110);
            label1.Margin = new Padding(2, 0, 2, 0);
            label1.Name = "label1";
            label1.Size = new Size(36, 15);
            label1.TabIndex = 9;
            label1.Text = "Wins:";
            // 
            // lblEstPower
            // 
            lblEstPower.AutoSize = true;
            lblEstPower.Location = new Point(57, 120);
            lblEstPower.Margin = new Padding(2, 0, 2, 0);
            lblEstPower.Name = "lblEstPower";
            lblEstPower.Size = new Size(61, 15);
            lblEstPower.TabIndex = 8;
            lblEstPower.Text = "Est. Power";
            // 
            // progressBar1
            // 
            progressBar1.BackColor = Color.Black;
            progressBar1.ForeColor = Color.Lime;
            progressBar1.Location = new Point(47, 137);
            progressBar1.Margin = new Padding(2);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(99, 22);
            progressBar1.Style = ProgressBarStyle.Marquee;
            progressBar1.TabIndex = 7;
            progressBar1.Value = 50;
            // 
            // lblExtraInfo
            // 
            lblExtraInfo.Location = new Point(6, 70);
            lblExtraInfo.Margin = new Padding(2, 0, 2, 0);
            lblExtraInfo.Name = "lblExtraInfo";
            lblExtraInfo.Size = new Size(319, 38);
            lblExtraInfo.TabIndex = 6;
            lblExtraInfo.Text = "Extra Info";
            //
            // lblWins
            //
            lblWins.Dock = DockStyle.Fill;
            lblWins.Font = new Font("Microsoft Sans Serif", 6.75F, FontStyle.Regular, GraphicsUnit.Point, 0);
            lblWins.Name = "lblWins";
            lblWins.TabIndex = 0;
            lblWins.TextAlign = ContentAlignment.MiddleCenter;
            //
            // lblLosses
            //
            lblLosses.Dock = DockStyle.Fill;
            lblLosses.Font = new Font("Microsoft Sans Serif", 6.75F, FontStyle.Regular, GraphicsUnit.Point, 0);
            lblLosses.Name = "lblLosses";
            lblLosses.TabIndex = 0;
            lblLosses.TextAlign = ContentAlignment.MiddleCenter;
            //
            // lblTies
            //
            lblTies.Dock = DockStyle.Fill;
            lblTies.Font = new Font("Microsoft Sans Serif", 6.75F, FontStyle.Regular, GraphicsUnit.Point, 0);
            lblTies.Name = "lblTies";
            lblTies.TabIndex = 0;
            lblTies.TextAlign = ContentAlignment.MiddleCenter;
            //
            // lblUbers
            //
            lblUbers.ForeColor = Color.Red;
            lblUbers.Location = new Point(266, 31);
            lblUbers.Name = "lblUbers";
            lblUbers.Size = new Size(62, 37);
            lblUbers.TabIndex = 16;
            lblUbers.TextAlign = ContentAlignment.MiddleCenter;
            //
            // tmrRandbat
            //
            tmrRandbat.Tick += tmrRandbat_Tick;
            // 
            // pk6
            // 
            pk6.Location = new Point(225, 31);
            pk6.Margin = new Padding(2);
            pk6.Name = "pk6";
            pk6.Size = new Size(37, 37);
            pk6.TabIndex = 5;
            pk6.TabStop = false;
            // 
            // pk5
            // 
            pk5.Location = new Point(183, 31);
            pk5.Margin = new Padding(2);
            pk5.Name = "pk5";
            pk5.Size = new Size(37, 37);
            pk5.TabIndex = 4;
            pk5.TabStop = false;
            // 
            // pk4
            // 
            pk4.Location = new Point(141, 31);
            pk4.Margin = new Padding(2);
            pk4.Name = "pk4";
            pk4.Size = new Size(37, 37);
            pk4.TabIndex = 3;
            pk4.TabStop = false;
            // 
            // pk3
            // 
            pk3.Location = new Point(99, 31);
            pk3.Margin = new Padding(2);
            pk3.Name = "pk3";
            pk3.Size = new Size(37, 37);
            pk3.TabIndex = 2;
            pk3.TabStop = false;
            // 
            // pk2
            // 
            pk2.Location = new Point(57, 31);
            pk2.Margin = new Padding(2);
            pk2.Name = "pk2";
            pk2.Size = new Size(37, 37);
            pk2.TabIndex = 1;
            pk2.TabStop = false;
            // 
            // pk1
            // 
            pk1.Location = new Point(15, 31);
            pk1.Margin = new Padding(2);
            pk1.Name = "pk1";
            pk1.Size = new Size(37, 37);
            pk1.TabIndex = 0;
            pk1.TabStop = false;
            // 
            // grpRulesModes
            // 
            grpRulesModes.Controls.Add(lblTerrain);
            grpRulesModes.Controls.Add(lblMode);
            grpRulesModes.Controls.Add(comboBox2);
            grpRulesModes.Controls.Add(comboBox1);
            grpRulesModes.Controls.Add(lblClauses);
            grpRulesModes.Controls.Add(listView1);
            grpRulesModes.Location = new Point(14, 219);
            grpRulesModes.Margin = new Padding(4, 3, 4, 3);
            grpRulesModes.Name = "grpRulesModes";
            grpRulesModes.Padding = new Padding(4, 3, 4, 3);
            grpRulesModes.Size = new Size(341, 188);
            grpRulesModes.TabIndex = 5;
            grpRulesModes.TabStop = false;
            grpRulesModes.Text = "Rules && Modes";
            // 
            // lblTerrain
            // 
            lblTerrain.AutoSize = true;
            lblTerrain.Location = new Point(211, 18);
            lblTerrain.Margin = new Padding(2, 0, 2, 0);
            lblTerrain.Name = "lblTerrain";
            lblTerrain.Size = new Size(42, 15);
            lblTerrain.TabIndex = 7;
            lblTerrain.Text = "Terrain";
            // 
            // lblMode
            // 
            lblMode.AutoSize = true;
            lblMode.Location = new Point(6, 18);
            lblMode.Margin = new Padding(2, 0, 2, 0);
            lblMode.Name = "lblMode";
            lblMode.Size = new Size(38, 15);
            lblMode.TabIndex = 6;
            lblMode.Text = "Mode";
            // 
            // comboBox2
            // 
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.FormattingEnabled = true;
            comboBox2.Location = new Point(215, 37);
            comboBox2.Margin = new Padding(4, 3, 4, 3);
            comboBox2.Name = "comboBox2";
            comboBox2.Size = new Size(114, 23);
            comboBox2.TabIndex = 5;
            // 
            // comboBox1
            // 
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox1.FormattingEnabled = true;
            comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            comboBox1.Location = new Point(7, 37);
            comboBox1.Margin = new Padding(4, 3, 4, 3);
            comboBox1.Name = "comboBox1";
            comboBox1.Size = new Size(190, 23);
            comboBox1.TabIndex = 4;
            // 
            // label4
            // 
            label4.Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label4.Location = new Point(28, 9);
            label4.Margin = new Padding(4, 0, 4, 0);
            label4.Name = "label4";
            label4.Size = new Size(310, 33);
            label4.TabIndex = 6;
            label4.Text = "You have been challenged!";
            // 
            // ChallengeWindow
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(354, 448);
            Controls.Add(label4);
            Controls.Add(grpRulesModes);
            Controls.Add(grpTeam);
            Controls.Add(btnCancel);
            Controls.Add(btnChallenge);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Margin = new Padding(2);
            Name = "ChallengeWindow";
            Text = "Player Details";
            FormClosing += ChallengeWindow_FormClosing;
            FormClosed += ChallengeWindow_FormClosed;
            Load += ChallengeWindow_Load;
            grpTeam.ResumeLayout(false);
            grpTeam.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            ((System.ComponentModel.ISupportInitialize)pk6).EndInit();
            ((System.ComponentModel.ISupportInitialize)pk5).EndInit();
            ((System.ComponentModel.ISupportInitialize)pk4).EndInit();
            ((System.ComponentModel.ISupportInitialize)pk3).EndInit();
            ((System.ComponentModel.ISupportInitialize)pk2).EndInit();
            ((System.ComponentModel.ISupportInitialize)pk1).EndInit();
            grpRulesModes.ResumeLayout(false);
            grpRulesModes.PerformLayout();
            ResumeLayout(false);
        }

        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;

        private System.Windows.Forms.Panel panel1;

        private System.Windows.Forms.Label label4;

        #endregion

        private System.Windows.Forms.Button btnChallenge;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.Label lblClauses;
        private System.Windows.Forms.GroupBox grpTeam;
        private System.Windows.Forms.PictureBox pk3;
        private System.Windows.Forms.PictureBox pk2;
        private System.Windows.Forms.PictureBox pk1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblEstPower;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblExtraInfo;
        private System.Windows.Forms.PictureBox pk6;
        private System.Windows.Forms.PictureBox pk5;
        private System.Windows.Forms.PictureBox pk4;
        private System.Windows.Forms.GroupBox grpRulesModes;
        private System.Windows.Forms.Label lblTerrain;
        private System.Windows.Forms.Label lblMode;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label lblWins;
        private System.Windows.Forms.Label lblLosses;
        private System.Windows.Forms.Label lblTies;
        private System.Windows.Forms.Label lblUbers;
        private System.Windows.Forms.ColumnHeader ruleColumn;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Timer tmrRandbat;
    }
}