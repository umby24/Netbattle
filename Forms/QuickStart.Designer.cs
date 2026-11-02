namespace Netbattle {
    partial class Form1 {
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnTeamBuilder = new System.Windows.Forms.Button();
            this.btnJoinServer = new System.Windows.Forms.Button();
            this.btnAbout = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnTeamBuilder
            // 
            this.btnTeamBuilder.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnTeamBuilder.BackgroundImage")));
            this.btnTeamBuilder.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnTeamBuilder.Location = new System.Drawing.Point(24, 15);
            this.btnTeamBuilder.Margin = new System.Windows.Forms.Padding(6);
            this.btnTeamBuilder.Name = "btnTeamBuilder";
            this.btnTeamBuilder.Size = new System.Drawing.Size(500, 75);
            this.btnTeamBuilder.TabIndex = 0;
            this.btnTeamBuilder.UseVisualStyleBackColor = true;
            this.btnTeamBuilder.Click += new System.EventHandler(this.btnTeamBuilder_Click);
            // 
            // btnJoinServer
            // 
            this.btnJoinServer.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnJoinServer.BackgroundImage")));
            this.btnJoinServer.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnJoinServer.Location = new System.Drawing.Point(25, 90);
            this.btnJoinServer.Margin = new System.Windows.Forms.Padding(6);
            this.btnJoinServer.Name = "btnJoinServer";
            this.btnJoinServer.Size = new System.Drawing.Size(500, 82);
            this.btnJoinServer.TabIndex = 1;
            this.btnJoinServer.UseVisualStyleBackColor = true;
            this.btnJoinServer.Click += new System.EventHandler(this.btnJoinServer_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnAbout.BackgroundImage")));
            this.btnAbout.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnAbout.Location = new System.Drawing.Point(24, 184);
            this.btnAbout.Margin = new System.Windows.Forms.Padding(6);
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Size = new System.Drawing.Size(187, 64);
            this.btnAbout.TabIndex = 2;
            this.btnAbout.UseVisualStyleBackColor = true;
            this.btnAbout.Click += new System.EventHandler(this.btnAbout_Click);
            // 
            // btnExit
            // 
            this.btnExit.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnExit.BackgroundImage")));
            this.btnExit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnExit.Location = new System.Drawing.Point(297, 184);
            this.btnExit.Margin = new System.Windows.Forms.Padding(6);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(194, 64);
            this.btnExit.TabIndex = 3;
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(548, 271);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnAbout);
            this.Controls.Add(this.btnJoinServer);
            this.Controls.Add(this.btnTeamBuilder);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "QuickStart";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.Button btnTeamBuilder;
        private System.Windows.Forms.Button btnJoinServer;
        private System.Windows.Forms.Button btnAbout;
        private System.Windows.Forms.Button btnExit;
    }
}

