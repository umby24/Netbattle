namespace NetbattleClient.Forms {
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
            btnTeamBuilder = new System.Windows.Forms.Button();
            btnJoinServer = new System.Windows.Forms.Button();
            btnAbout = new System.Windows.Forms.Button();
            btnExit = new System.Windows.Forms.Button();
            menuStrip1 = new System.Windows.Forms.MenuStrip();
            fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            newToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // btnTeamBuilder
            // 
            btnTeamBuilder.BackgroundImage = global::NetbattleClient.Properties.Resources.TeamBuilder_Up;
            btnTeamBuilder.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            btnTeamBuilder.FlatAppearance.BorderSize = 0;
            btnTeamBuilder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            btnTeamBuilder.Location = new System.Drawing.Point(11, 9);
            btnTeamBuilder.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btnTeamBuilder.Name = "btnTeamBuilder";
            btnTeamBuilder.Size = new System.Drawing.Size(250, 40);
            btnTeamBuilder.TabIndex = 0;
            btnTeamBuilder.UseVisualStyleBackColor = true;
            btnTeamBuilder.Click += btnTeamBuilder_Click;
            // 
            // btnJoinServer
            // 
            btnJoinServer.BackgroundImage = global::NetbattleClient.Properties.Resources.StartBattling_Up;
            btnJoinServer.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            btnJoinServer.FlatAppearance.BorderSize = 0;
            btnJoinServer.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            btnJoinServer.Location = new System.Drawing.Point(12, 55);
            btnJoinServer.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btnJoinServer.Name = "btnJoinServer";
            btnJoinServer.Size = new System.Drawing.Size(250, 40);
            btnJoinServer.TabIndex = 1;
            btnJoinServer.UseVisualStyleBackColor = true;
            btnJoinServer.Click += btnJoinServer_Click;
            // 
            // btnAbout
            // 
            btnAbout.BackgroundImage = global::NetbattleClient.Properties.Resources.About_Up;
            btnAbout.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            btnAbout.FlatAppearance.BorderSize = 0;
            btnAbout.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            btnAbout.Location = new System.Drawing.Point(11, 102);
            btnAbout.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btnAbout.Name = "btnAbout";
            btnAbout.Size = new System.Drawing.Size(125, 40);
            btnAbout.TabIndex = 2;
            btnAbout.UseVisualStyleBackColor = true;
            btnAbout.Click += btnAbout_Click;
            // 
            // btnExit
            // 
            btnExit.BackgroundImage = global::NetbattleClient.Properties.Resources.Exit_Up;
            btnExit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            btnExit.FlatAppearance.BorderSize = 0;
            btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            btnExit.Location = new System.Drawing.Point(144, 102);
            btnExit.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btnExit.Name = "btnExit";
            btnExit.Size = new System.Drawing.Size(125, 40);
            btnExit.TabIndex = 3;
            btnExit.UseVisualStyleBackColor = true;
            btnExit.Click += btnExit_Click;
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { fileToolStripMenuItem });
            menuStrip1.Location = new System.Drawing.Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new System.Drawing.Size(271, 24);
            menuStrip1.TabIndex = 4;
            menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { newToolStripMenuItem, openToolStripMenuItem, settingsToolStripMenuItem });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            fileToolStripMenuItem.Text = "&File";
            // 
            // newToolStripMenuItem
            // 
            newToolStripMenuItem.Name = "newToolStripMenuItem";
            newToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            newToolStripMenuItem.Text = "&New";
            // 
            // openToolStripMenuItem
            // 
            openToolStripMenuItem.Name = "openToolStripMenuItem";
            openToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            openToolStripMenuItem.Text = "&Open";
            // 
            // settingsToolStripMenuItem
            // 
            settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            settingsToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            settingsToolStripMenuItem.Text = "&Settings";
            settingsToolStripMenuItem.Click += settingsToolStripMenuItem_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            BackgroundImage = ((System.Drawing.Image)resources.GetObject("$this.BackgroundImage"));
            BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            ClientSize = new System.Drawing.Size(271, 160);
            Controls.Add(btnExit);
            Controls.Add(btnAbout);
            Controls.Add(btnJoinServer);
            Controls.Add(btnTeamBuilder);
            Controls.Add(menuStrip1);
            Icon = ((System.Drawing.Icon)resources.GetObject("$this.Icon"));
            MainMenuStrip = menuStrip1;
            Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "QuickStart";
            Load += Form1_Load;
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem newToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;

        #endregion

        private System.Windows.Forms.Button btnTeamBuilder;
        private System.Windows.Forms.Button btnJoinServer;
        private System.Windows.Forms.Button btnAbout;
        private System.Windows.Forms.Button btnExit;
    }
}

