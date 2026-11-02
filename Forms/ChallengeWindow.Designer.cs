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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChallengeWindow));
            this.btnChallenge = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.listView1 = new System.Windows.Forms.ListView();
            this.lblClauses = new System.Windows.Forms.Label();
            this.grpTeam = new System.Windows.Forms.GroupBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lblEstPower = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblExtraInfo = new System.Windows.Forms.Label();
            this.pk6 = new System.Windows.Forms.PictureBox();
            this.pk5 = new System.Windows.Forms.PictureBox();
            this.pk4 = new System.Windows.Forms.PictureBox();
            this.pk3 = new System.Windows.Forms.PictureBox();
            this.pk2 = new System.Windows.Forms.PictureBox();
            this.pk1 = new System.Windows.Forms.PictureBox();
            this.grpRulesModes = new System.Windows.Forms.GroupBox();
            this.lblTerrain = new System.Windows.Forms.Label();
            this.lblMode = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.grpTeam.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk1)).BeginInit();
            this.grpRulesModes.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnChallenge
            // 
            this.btnChallenge.Location = new System.Drawing.Point(11, 358);
            this.btnChallenge.Margin = new System.Windows.Forms.Padding(2);
            this.btnChallenge.Name = "btnChallenge";
            this.btnChallenge.Size = new System.Drawing.Size(142, 19);
            this.btnChallenge.TabIndex = 0;
            this.btnChallenge.Text = "Challenge!";
            this.btnChallenge.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(157, 358);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(2);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(135, 19);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // listView1
            // 
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(5, 71);
            this.listView1.Margin = new System.Windows.Forms.Padding(2);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(278, 80);
            this.listView1.TabIndex = 2;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.SelectedIndexChanged += new System.EventHandler(this.listView1_SelectedIndexChanged);
            // 
            // lblClauses
            // 
            this.lblClauses.AutoSize = true;
            this.lblClauses.Location = new System.Drawing.Point(5, 56);
            this.lblClauses.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblClauses.Name = "lblClauses";
            this.lblClauses.Size = new System.Drawing.Size(44, 13);
            this.lblClauses.TabIndex = 3;
            this.lblClauses.Text = "Clauses";
            // 
            // grpTeam
            // 
            this.grpTeam.Controls.Add(this.panel3);
            this.grpTeam.Controls.Add(this.panel2);
            this.grpTeam.Controls.Add(this.panel1);
            this.grpTeam.Controls.Add(this.pictureBox1);
            this.grpTeam.Controls.Add(this.label3);
            this.grpTeam.Controls.Add(this.label2);
            this.grpTeam.Controls.Add(this.label1);
            this.grpTeam.Controls.Add(this.lblEstPower);
            this.grpTeam.Controls.Add(this.progressBar1);
            this.grpTeam.Controls.Add(this.lblExtraInfo);
            this.grpTeam.Controls.Add(this.pk6);
            this.grpTeam.Controls.Add(this.pk5);
            this.grpTeam.Controls.Add(this.pk4);
            this.grpTeam.Controls.Add(this.pk3);
            this.grpTeam.Controls.Add(this.pk2);
            this.grpTeam.Controls.Add(this.pk1);
            this.grpTeam.Location = new System.Drawing.Point(11, 39);
            this.grpTeam.Margin = new System.Windows.Forms.Padding(2);
            this.grpTeam.Name = "grpTeam";
            this.grpTeam.Padding = new System.Windows.Forms.Padding(2);
            this.grpTeam.Size = new System.Drawing.Size(284, 146);
            this.grpTeam.TabIndex = 4;
            this.grpTeam.TabStop = false;
            this.grpTeam.Text = "Player and Team Info";
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel3.Location = new System.Drawing.Point(235, 111);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(33, 33);
            this.panel3.TabIndex = 15;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Location = new System.Drawing.Point(196, 111);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(33, 33);
            this.panel2.TabIndex = 14;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Location = new System.Drawing.Point(157, 111);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(33, 33);
            this.panel1.TabIndex = 13;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(4, 106);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(32, 32);
            this.pictureBox1.TabIndex = 12;
            this.pictureBox1.TabStop = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(238, 95);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(30, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "Ties:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(196, 95);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "Losses:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(156, 95);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(34, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Wins:";
            // 
            // lblEstPower
            // 
            this.lblEstPower.AutoSize = true;
            this.lblEstPower.Location = new System.Drawing.Point(49, 104);
            this.lblEstPower.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblEstPower.Name = "lblEstPower";
            this.lblEstPower.Size = new System.Drawing.Size(58, 13);
            this.lblEstPower.TabIndex = 8;
            this.lblEstPower.Text = "Est. Power";
            // 
            // progressBar1
            // 
            this.progressBar1.BackColor = System.Drawing.Color.Black;
            this.progressBar1.ForeColor = System.Drawing.Color.Lime;
            this.progressBar1.Location = new System.Drawing.Point(40, 119);
            this.progressBar1.Margin = new System.Windows.Forms.Padding(2);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(85, 19);
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar1.TabIndex = 7;
            this.progressBar1.Value = 50;
            // 
            // lblExtraInfo
            // 
            this.lblExtraInfo.AutoSize = true;
            this.lblExtraInfo.Location = new System.Drawing.Point(14, 61);
            this.lblExtraInfo.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblExtraInfo.Name = "lblExtraInfo";
            this.lblExtraInfo.Size = new System.Drawing.Size(44, 13);
            this.lblExtraInfo.TabIndex = 6;
            this.lblExtraInfo.Text = "Clauses";
            // 
            // pk6
            // 
            this.pk6.Location = new System.Drawing.Point(193, 27);
            this.pk6.Margin = new System.Windows.Forms.Padding(2);
            this.pk6.Name = "pk6";
            this.pk6.Size = new System.Drawing.Size(32, 32);
            this.pk6.TabIndex = 5;
            this.pk6.TabStop = false;
            // 
            // pk5
            // 
            this.pk5.Location = new System.Drawing.Point(157, 27);
            this.pk5.Margin = new System.Windows.Forms.Padding(2);
            this.pk5.Name = "pk5";
            this.pk5.Size = new System.Drawing.Size(32, 32);
            this.pk5.TabIndex = 4;
            this.pk5.TabStop = false;
            // 
            // pk4
            // 
            this.pk4.Location = new System.Drawing.Point(121, 27);
            this.pk4.Margin = new System.Windows.Forms.Padding(2);
            this.pk4.Name = "pk4";
            this.pk4.Size = new System.Drawing.Size(32, 32);
            this.pk4.TabIndex = 3;
            this.pk4.TabStop = false;
            // 
            // pk3
            // 
            this.pk3.Location = new System.Drawing.Point(85, 27);
            this.pk3.Margin = new System.Windows.Forms.Padding(2);
            this.pk3.Name = "pk3";
            this.pk3.Size = new System.Drawing.Size(32, 32);
            this.pk3.TabIndex = 2;
            this.pk3.TabStop = false;
            this.pk3.Click += new System.EventHandler(this.pictureBox3_Click);
            // 
            // pk2
            // 
            this.pk2.Location = new System.Drawing.Point(49, 27);
            this.pk2.Margin = new System.Windows.Forms.Padding(2);
            this.pk2.Name = "pk2";
            this.pk2.Size = new System.Drawing.Size(32, 32);
            this.pk2.TabIndex = 1;
            this.pk2.TabStop = false;
            // 
            // pk1
            // 
            this.pk1.Location = new System.Drawing.Point(13, 27);
            this.pk1.Margin = new System.Windows.Forms.Padding(2);
            this.pk1.Name = "pk1";
            this.pk1.Size = new System.Drawing.Size(32, 32);
            this.pk1.TabIndex = 0;
            this.pk1.TabStop = false;
            // 
            // grpRulesModes
            // 
            this.grpRulesModes.Controls.Add(this.lblTerrain);
            this.grpRulesModes.Controls.Add(this.lblMode);
            this.grpRulesModes.Controls.Add(this.comboBox2);
            this.grpRulesModes.Controls.Add(this.comboBox1);
            this.grpRulesModes.Controls.Add(this.lblClauses);
            this.grpRulesModes.Controls.Add(this.listView1);
            this.grpRulesModes.Location = new System.Drawing.Point(12, 190);
            this.grpRulesModes.Name = "grpRulesModes";
            this.grpRulesModes.Size = new System.Drawing.Size(292, 163);
            this.grpRulesModes.TabIndex = 5;
            this.grpRulesModes.TabStop = false;
            this.grpRulesModes.Text = "Rules && Modes";
            // 
            // lblTerrain
            // 
            this.lblTerrain.AutoSize = true;
            this.lblTerrain.Location = new System.Drawing.Point(181, 16);
            this.lblTerrain.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblTerrain.Name = "lblTerrain";
            this.lblTerrain.Size = new System.Drawing.Size(40, 13);
            this.lblTerrain.TabIndex = 7;
            this.lblTerrain.Text = "Terrain";
            // 
            // lblMode
            // 
            this.lblMode.AutoSize = true;
            this.lblMode.Location = new System.Drawing.Point(5, 16);
            this.lblMode.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblMode.Name = "lblMode";
            this.lblMode.Size = new System.Drawing.Size(34, 13);
            this.lblMode.TabIndex = 6;
            this.lblMode.Text = "Mode";
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(184, 32);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(98, 21);
            this.comboBox2.TabIndex = 5;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(6, 32);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(163, 21);
            this.comboBox1.TabIndex = 4;
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(24, 8);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(266, 29);
            this.label4.TabIndex = 6;
            this.label4.Text = "You have been challenged!";
            // 
            // ChallengeWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(303, 388);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.grpRulesModes);
            this.Controls.Add(this.grpTeam);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnChallenge);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "ChallengeWindow";
            this.Text = "Player Details";
            this.grpTeam.ResumeLayout(false);
            this.grpTeam.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk1)).EndInit();
            this.grpRulesModes.ResumeLayout(false);
            this.grpRulesModes.PerformLayout();
            this.ResumeLayout(false);
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
    }
}