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
        private void InitializeComponent() {
            this.btnChallenge = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.listView1 = new System.Windows.Forms.ListView();
            this.lblClauses = new System.Windows.Forms.Label();
            this.grpTeam = new System.Windows.Forms.GroupBox();
            this.pk1 = new System.Windows.Forms.PictureBox();
            this.pk2 = new System.Windows.Forms.PictureBox();
            this.pk3 = new System.Windows.Forms.PictureBox();
            this.pk4 = new System.Windows.Forms.PictureBox();
            this.pk5 = new System.Windows.Forms.PictureBox();
            this.pk6 = new System.Windows.Forms.PictureBox();
            this.lblExtraInfo = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblEstPower = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.grpTeam.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pk1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk6)).BeginInit();
            this.SuspendLayout();
            // 
            // btnChallenge
            // 
            this.btnChallenge.Location = new System.Drawing.Point(12, 455);
            this.btnChallenge.Name = "btnChallenge";
            this.btnChallenge.Size = new System.Drawing.Size(225, 23);
            this.btnChallenge.TabIndex = 0;
            this.btnChallenge.Text = "Challenge!";
            this.btnChallenge.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(260, 455);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(225, 23);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // listView1
            // 
            this.listView1.Location = new System.Drawing.Point(12, 352);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(473, 97);
            this.listView1.TabIndex = 2;
            this.listView1.UseCompatibleStateImageBehavior = false;
            // 
            // lblClauses
            // 
            this.lblClauses.AutoSize = true;
            this.lblClauses.Location = new System.Drawing.Point(12, 332);
            this.lblClauses.Name = "lblClauses";
            this.lblClauses.Size = new System.Drawing.Size(58, 17);
            this.lblClauses.TabIndex = 3;
            this.lblClauses.Text = "Clauses";
            // 
            // grpTeam
            // 
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
            this.grpTeam.Location = new System.Drawing.Point(0, 0);
            this.grpTeam.Name = "grpTeam";
            this.grpTeam.Size = new System.Drawing.Size(485, 180);
            this.grpTeam.TabIndex = 4;
            this.grpTeam.TabStop = false;
            this.grpTeam.Text = "Pokemon and Team Info";
            // 
            // pk1
            // 
            this.pk1.Location = new System.Drawing.Point(17, 33);
            this.pk1.Name = "pk1";
            this.pk1.Size = new System.Drawing.Size(53, 50);
            this.pk1.TabIndex = 0;
            this.pk1.TabStop = false;
            // 
            // pk2
            // 
            this.pk2.Location = new System.Drawing.Point(89, 33);
            this.pk2.Name = "pk2";
            this.pk2.Size = new System.Drawing.Size(53, 50);
            this.pk2.TabIndex = 1;
            this.pk2.TabStop = false;
            // 
            // pk3
            // 
            this.pk3.Location = new System.Drawing.Point(163, 33);
            this.pk3.Name = "pk3";
            this.pk3.Size = new System.Drawing.Size(53, 50);
            this.pk3.TabIndex = 2;
            this.pk3.TabStop = false;
            this.pk3.Click += new System.EventHandler(this.pictureBox3_Click);
            // 
            // pk4
            // 
            this.pk4.Location = new System.Drawing.Point(247, 33);
            this.pk4.Name = "pk4";
            this.pk4.Size = new System.Drawing.Size(53, 50);
            this.pk4.TabIndex = 3;
            this.pk4.TabStop = false;
            // 
            // pk5
            // 
            this.pk5.Location = new System.Drawing.Point(322, 33);
            this.pk5.Name = "pk5";
            this.pk5.Size = new System.Drawing.Size(53, 50);
            this.pk5.TabIndex = 4;
            this.pk5.TabStop = false;
            // 
            // pk6
            // 
            this.pk6.Location = new System.Drawing.Point(397, 33);
            this.pk6.Name = "pk6";
            this.pk6.Size = new System.Drawing.Size(53, 50);
            this.pk6.TabIndex = 5;
            this.pk6.TabStop = false;
            // 
            // lblExtraInfo
            // 
            this.lblExtraInfo.AutoSize = true;
            this.lblExtraInfo.Location = new System.Drawing.Point(14, 97);
            this.lblExtraInfo.Name = "lblExtraInfo";
            this.lblExtraInfo.Size = new System.Drawing.Size(58, 17);
            this.lblExtraInfo.TabIndex = 6;
            this.lblExtraInfo.Text = "Clauses";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(17, 151);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(100, 23);
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar1.TabIndex = 7;
            // 
            // lblEstPower
            // 
            this.lblEstPower.AutoSize = true;
            this.lblEstPower.Location = new System.Drawing.Point(14, 131);
            this.lblEstPower.Name = "lblEstPower";
            this.lblEstPower.Size = new System.Drawing.Size(75, 17);
            this.lblEstPower.TabIndex = 8;
            this.lblEstPower.Text = "Est. Power";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(178, 151);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 17);
            this.label1.TabIndex = 9;
            this.label1.Text = "Wins:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(303, 151);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(57, 17);
            this.label2.TabIndex = 10;
            this.label2.Text = "Losses:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(303, 131);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(39, 17);
            this.label3.TabIndex = 11;
            this.label3.Text = "Ties:";
            // 
            // ChallengeWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(497, 490);
            this.Controls.Add(this.grpTeam);
            this.Controls.Add(this.lblClauses);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnChallenge);
            this.Name = "ChallengeWindow";
            this.Text = "ChallengeWindow";
            this.grpTeam.ResumeLayout(false);
            this.grpTeam.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pk1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pk6)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

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
    }
}