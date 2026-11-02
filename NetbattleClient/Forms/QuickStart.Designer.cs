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
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            btnTeamBuilder = new Button();
            btnJoinServer = new Button();
            btnAbout = new Button();
            btnExit = new Button();
            SuspendLayout();
            // 
            // btnTeamBuilder
            // 
            btnTeamBuilder.BackgroundImage = Properties.Resources.TeamBuilder_Up;
            btnTeamBuilder.BackgroundImageLayout = ImageLayout.Stretch;
            btnTeamBuilder.FlatAppearance.BorderSize = 0;
            btnTeamBuilder.FlatStyle = FlatStyle.Flat;
            btnTeamBuilder.Location = new Point(11, 9);
            btnTeamBuilder.Margin = new Padding(4, 3, 4, 3);
            btnTeamBuilder.Name = "btnTeamBuilder";
            btnTeamBuilder.Size = new Size(250, 40);
            btnTeamBuilder.TabIndex = 0;
            btnTeamBuilder.UseVisualStyleBackColor = true;
            btnTeamBuilder.Click += btnTeamBuilder_Click;
            // 
            // btnJoinServer
            // 
            btnJoinServer.BackgroundImage = Properties.Resources.StartBattling_Up;
            btnJoinServer.BackgroundImageLayout = ImageLayout.Center;
            btnJoinServer.FlatAppearance.BorderSize = 0;
            btnJoinServer.FlatStyle = FlatStyle.Flat;
            btnJoinServer.Location = new Point(12, 55);
            btnJoinServer.Margin = new Padding(4, 3, 4, 3);
            btnJoinServer.Name = "btnJoinServer";
            btnJoinServer.Size = new Size(250, 40);
            btnJoinServer.TabIndex = 1;
            btnJoinServer.UseVisualStyleBackColor = true;
            btnJoinServer.Click += btnJoinServer_Click;
            // 
            // btnAbout
            // 
            btnAbout.BackgroundImage = Properties.Resources.About_Up;
            btnAbout.BackgroundImageLayout = ImageLayout.Stretch;
            btnAbout.FlatAppearance.BorderSize = 0;
            btnAbout.FlatStyle = FlatStyle.Flat;
            btnAbout.Location = new Point(11, 102);
            btnAbout.Margin = new Padding(4, 3, 4, 3);
            btnAbout.Name = "btnAbout";
            btnAbout.Size = new Size(125, 40);
            btnAbout.TabIndex = 2;
            btnAbout.UseVisualStyleBackColor = true;
            btnAbout.Click += btnAbout_Click;
            // 
            // btnExit
            // 
            btnExit.BackgroundImage = Properties.Resources.Exit_Up;
            btnExit.BackgroundImageLayout = ImageLayout.Stretch;
            btnExit.FlatAppearance.BorderSize = 0;
            btnExit.FlatStyle = FlatStyle.Flat;
            btnExit.Location = new Point(144, 102);
            btnExit.Margin = new Padding(4, 3, 4, 3);
            btnExit.Name = "btnExit";
            btnExit.Size = new Size(125, 40);
            btnExit.TabIndex = 3;
            btnExit.UseVisualStyleBackColor = true;
            btnExit.Click += btnExit_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackgroundImage = (Image)resources.GetObject("$this.BackgroundImage");
            BackgroundImageLayout = ImageLayout.Stretch;
            ClientSize = new Size(271, 160);
            Controls.Add(btnExit);
            Controls.Add(btnAbout);
            Controls.Add(btnJoinServer);
            Controls.Add(btnTeamBuilder);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Margin = new Padding(4, 3, 4, 3);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "QuickStart";
            Load += Form1_Load;
            ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.Button btnTeamBuilder;
        private System.Windows.Forms.Button btnJoinServer;
        private System.Windows.Forms.Button btnAbout;
        private System.Windows.Forms.Button btnExit;
    }
}

