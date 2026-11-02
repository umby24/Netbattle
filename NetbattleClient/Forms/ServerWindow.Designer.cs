namespace Netbattle.Forms {
    partial class ServerWindow {
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
            components = new System.ComponentModel.Container();
            System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem("umby24", 3);
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ServerWindow));
            lstPlayers = new System.Windows.Forms.ListView();
            nameColumn = new System.Windows.Forms.ColumnHeader();
            playerMenu = new System.Windows.Forms.ContextMenuStrip(components);
            challengeInfoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            watchBattleToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            privateMessageToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            controlWindowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            kickToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            imageList1 = new System.Windows.Forms.ImageList(components);
            txtInput = new System.Windows.Forms.TextBox();
            btnSend = new System.Windows.Forms.Button();
            menuStrip1 = new System.Windows.Forms.MenuStrip();
            fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            saveLogToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            toolStripMenuItem2 = new System.Windows.Forms.ToolStripSeparator();
            exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            teamToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            rearrangeTeamToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            accessBoxToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            changeItemsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            openTeamBuilderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            loadAnotherTeamToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            playerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            challengeInfoToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            watchBattleToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            privateMessageToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            controlWindowToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            toolStripMenuItem3 = new System.Windows.Forms.ToolStripSeparator();
            kickToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            soundToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            musicToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            updatePlayerSpeedsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            allowMultilinePasteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            showVersionIconsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            colorNamesInListToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            toolStripMenuItem4 = new System.Windows.Forms.ToolStripSeparator();
            awayToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            windowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            txtChat = new System.Windows.Forms.RichTextBox();
            playerMenu.SuspendLayout();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // lstPlayers
            // 
            lstPlayers.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] { nameColumn });
            lstPlayers.ContextMenuStrip = playerMenu;
            lstPlayers.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)0));
            lstPlayers.FullRowSelect = true;
            lstPlayers.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            lstPlayers.Items.AddRange(new System.Windows.Forms.ListViewItem[] { listViewItem1 });
            lstPlayers.LargeImageList = imageList1;
            lstPlayers.Location = new System.Drawing.Point(0, 60);
            lstPlayers.Margin = new System.Windows.Forms.Padding(4);
            lstPlayers.MultiSelect = false;
            lstPlayers.Name = "lstPlayers";
            lstPlayers.Size = new System.Drawing.Size(254, 727);
            lstPlayers.SmallImageList = imageList1;
            lstPlayers.TabIndex = 1;
            lstPlayers.TileSize = new System.Drawing.Size(10, 50);
            lstPlayers.UseCompatibleStateImageBehavior = false;
            lstPlayers.View = System.Windows.Forms.View.List;
            // 
            // nameColumn
            // 
            nameColumn.Name = "nameColumn";
            nameColumn.Text = "Name";
            nameColumn.Width = 115;
            // 
            // playerMenu
            // 
            playerMenu.ImageScalingSize = new System.Drawing.Size(32, 32);
            playerMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { challengeInfoToolStripMenuItem, watchBattleToolStripMenuItem, privateMessageToolStripMenuItem, controlWindowToolStripMenuItem, toolStripMenuItem1, kickToolStripMenuItem });
            playerMenu.Name = "playerMenu";
            playerMenu.Size = new System.Drawing.Size(262, 200);
            // 
            // challengeInfoToolStripMenuItem
            // 
            challengeInfoToolStripMenuItem.Name = "challengeInfoToolStripMenuItem";
            challengeInfoToolStripMenuItem.Size = new System.Drawing.Size(261, 38);
            challengeInfoToolStripMenuItem.Text = "Challenge/Info";
            // 
            // watchBattleToolStripMenuItem
            // 
            watchBattleToolStripMenuItem.Name = "watchBattleToolStripMenuItem";
            watchBattleToolStripMenuItem.Size = new System.Drawing.Size(261, 38);
            watchBattleToolStripMenuItem.Text = "Watch Battle";
            // 
            // privateMessageToolStripMenuItem
            // 
            privateMessageToolStripMenuItem.Name = "privateMessageToolStripMenuItem";
            privateMessageToolStripMenuItem.Size = new System.Drawing.Size(261, 38);
            privateMessageToolStripMenuItem.Text = "Private Message";
            privateMessageToolStripMenuItem.Click += privateMessageToolStripMenuItem_Click;
            // 
            // controlWindowToolStripMenuItem
            // 
            controlWindowToolStripMenuItem.Name = "controlWindowToolStripMenuItem";
            controlWindowToolStripMenuItem.Size = new System.Drawing.Size(261, 38);
            controlWindowToolStripMenuItem.Text = "Control Window";
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new System.Drawing.Size(258, 6);
            // 
            // kickToolStripMenuItem
            // 
            kickToolStripMenuItem.Name = "kickToolStripMenuItem";
            kickToolStripMenuItem.Size = new System.Drawing.Size(261, 38);
            kickToolStripMenuItem.Text = "&Kick";
            // 
            // imageList1
            // 
            imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)resources.GetObject("imageList1.ImageStream"));
            imageList1.TransparentColor = System.Drawing.Color.Transparent;
            imageList1.Images.SetKeyName(0, "trainer1.ico");
            imageList1.Images.SetKeyName(1, "trainer2.ico");
            imageList1.Images.SetKeyName(2, "trainer3.ico");
            imageList1.Images.SetKeyName(3, "trainer4.ico");
            imageList1.Images.SetKeyName(4, "trainer5.ico");
            imageList1.Images.SetKeyName(5, "trainer6.ico");
            imageList1.Images.SetKeyName(6, "trainer7.ico");
            imageList1.Images.SetKeyName(7, "trainer8.ico");
            imageList1.Images.SetKeyName(8, "trainer9.ico");
            imageList1.Images.SetKeyName(9, "trainer10.ico");
            imageList1.Images.SetKeyName(10, "trainer11.ico");
            imageList1.Images.SetKeyName(11, "trainer12.ico");
            imageList1.Images.SetKeyName(12, "trainer13.ico");
            imageList1.Images.SetKeyName(13, "trainer14.ico");
            imageList1.Images.SetKeyName(14, "trainer15.ico");
            // 
            // txtInput
            // 
            txtInput.Location = new System.Drawing.Point(266, 748);
            txtInput.Margin = new System.Windows.Forms.Padding(6);
            txtInput.MaxLength = 200;
            txtInput.Name = "txtInput";
            txtInput.Size = new System.Drawing.Size(1250, 39);
            txtInput.TabIndex = 2;
            // 
            // btnSend
            // 
            btnSend.Location = new System.Drawing.Point(1538, 748);
            btnSend.Margin = new System.Windows.Forms.Padding(6);
            btnSend.Name = "btnSend";
            btnSend.Size = new System.Drawing.Size(150, 39);
            btnSend.TabIndex = 3;
            btnSend.Text = "Send";
            btnSend.UseVisualStyleBackColor = true;
            btnSend.Click += btnSend_Click;
            // 
            // menuStrip1
            // 
            menuStrip1.ImageScalingSize = new System.Drawing.Size(32, 32);
            menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { fileToolStripMenuItem, teamToolStripMenuItem, playerToolStripMenuItem, optionsToolStripMenuItem, windowToolStripMenuItem });
            menuStrip1.Location = new System.Drawing.Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Padding = new System.Windows.Forms.Padding(12, 4, 0, 4);
            menuStrip1.Size = new System.Drawing.Size(1696, 44);
            menuStrip1.TabIndex = 4;
            menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { saveLogToolStripMenuItem, toolStripMenuItem2, exitToolStripMenuItem });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new System.Drawing.Size(71, 36);
            fileToolStripMenuItem.Text = "&File";
            // 
            // saveLogToolStripMenuItem
            // 
            saveLogToolStripMenuItem.Name = "saveLogToolStripMenuItem";
            saveLogToolStripMenuItem.Size = new System.Drawing.Size(258, 44);
            saveLogToolStripMenuItem.Text = "&Save Log...";
            // 
            // toolStripMenuItem2
            // 
            toolStripMenuItem2.Name = "toolStripMenuItem2";
            toolStripMenuItem2.Size = new System.Drawing.Size(255, 6);
            // 
            // exitToolStripMenuItem
            // 
            exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            exitToolStripMenuItem.Size = new System.Drawing.Size(258, 44);
            exitToolStripMenuItem.Text = "E&xit";
            // 
            // teamToolStripMenuItem
            // 
            teamToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { rearrangeTeamToolStripMenuItem, accessBoxToolStripMenuItem, changeItemsToolStripMenuItem, openTeamBuilderToolStripMenuItem, loadAnotherTeamToolStripMenuItem });
            teamToolStripMenuItem.Name = "teamToolStripMenuItem";
            teamToolStripMenuItem.Size = new System.Drawing.Size(91, 36);
            teamToolStripMenuItem.Text = "&Team";
            // 
            // rearrangeTeamToolStripMenuItem
            // 
            rearrangeTeamToolStripMenuItem.Name = "rearrangeTeamToolStripMenuItem";
            rearrangeTeamToolStripMenuItem.Size = new System.Drawing.Size(355, 44);
            rearrangeTeamToolStripMenuItem.Text = "Rearrange Team";
            // 
            // accessBoxToolStripMenuItem
            // 
            accessBoxToolStripMenuItem.Name = "accessBoxToolStripMenuItem";
            accessBoxToolStripMenuItem.Size = new System.Drawing.Size(355, 44);
            accessBoxToolStripMenuItem.Text = "Access Box";
            // 
            // changeItemsToolStripMenuItem
            // 
            changeItemsToolStripMenuItem.Name = "changeItemsToolStripMenuItem";
            changeItemsToolStripMenuItem.Size = new System.Drawing.Size(355, 44);
            changeItemsToolStripMenuItem.Text = "Change Items";
            // 
            // openTeamBuilderToolStripMenuItem
            // 
            openTeamBuilderToolStripMenuItem.Name = "openTeamBuilderToolStripMenuItem";
            openTeamBuilderToolStripMenuItem.Size = new System.Drawing.Size(355, 44);
            openTeamBuilderToolStripMenuItem.Text = "&Open Team Builder";
            // 
            // loadAnotherTeamToolStripMenuItem
            // 
            loadAnotherTeamToolStripMenuItem.Name = "loadAnotherTeamToolStripMenuItem";
            loadAnotherTeamToolStripMenuItem.Size = new System.Drawing.Size(355, 44);
            loadAnotherTeamToolStripMenuItem.Text = "Load Another Team";
            // 
            // playerToolStripMenuItem
            // 
            playerToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { challengeInfoToolStripMenuItem1, watchBattleToolStripMenuItem1, privateMessageToolStripMenuItem1, controlWindowToolStripMenuItem1, toolStripMenuItem3, kickToolStripMenuItem1 });
            playerToolStripMenuItem.Name = "playerToolStripMenuItem";
            playerToolStripMenuItem.Size = new System.Drawing.Size(98, 36);
            playerToolStripMenuItem.Text = "&Player";
            // 
            // challengeInfoToolStripMenuItem1
            // 
            challengeInfoToolStripMenuItem1.Name = "challengeInfoToolStripMenuItem1";
            challengeInfoToolStripMenuItem1.Size = new System.Drawing.Size(335, 44);
            challengeInfoToolStripMenuItem1.Text = "&Challenge/Info...";
            // 
            // watchBattleToolStripMenuItem1
            // 
            watchBattleToolStripMenuItem1.Name = "watchBattleToolStripMenuItem1";
            watchBattleToolStripMenuItem1.Size = new System.Drawing.Size(335, 44);
            watchBattleToolStripMenuItem1.Text = "&Watch Battle...";
            // 
            // privateMessageToolStripMenuItem1
            // 
            privateMessageToolStripMenuItem1.Name = "privateMessageToolStripMenuItem1";
            privateMessageToolStripMenuItem1.Size = new System.Drawing.Size(335, 44);
            privateMessageToolStripMenuItem1.Text = "&Private Message...";
            // 
            // controlWindowToolStripMenuItem1
            // 
            controlWindowToolStripMenuItem1.Name = "controlWindowToolStripMenuItem1";
            controlWindowToolStripMenuItem1.Size = new System.Drawing.Size(335, 44);
            controlWindowToolStripMenuItem1.Text = "C&ontrol Window";
            // 
            // toolStripMenuItem3
            // 
            toolStripMenuItem3.Name = "toolStripMenuItem3";
            toolStripMenuItem3.Size = new System.Drawing.Size(332, 6);
            // 
            // kickToolStripMenuItem1
            // 
            kickToolStripMenuItem1.Name = "kickToolStripMenuItem1";
            kickToolStripMenuItem1.Size = new System.Drawing.Size(335, 44);
            kickToolStripMenuItem1.Text = "&Kick";
            // 
            // optionsToolStripMenuItem
            // 
            optionsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { soundToolStripMenuItem, musicToolStripMenuItem, updatePlayerSpeedsToolStripMenuItem, allowMultilinePasteToolStripMenuItem, showVersionIconsToolStripMenuItem, colorNamesInListToolStripMenuItem, toolStripMenuItem4, awayToolStripMenuItem });
            optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            optionsToolStripMenuItem.Size = new System.Drawing.Size(118, 36);
            optionsToolStripMenuItem.Text = "&Options";
            // 
            // soundToolStripMenuItem
            // 
            soundToolStripMenuItem.Name = "soundToolStripMenuItem";
            soundToolStripMenuItem.Size = new System.Drawing.Size(379, 44);
            soundToolStripMenuItem.Text = "&Sound";
            // 
            // musicToolStripMenuItem
            // 
            musicToolStripMenuItem.Name = "musicToolStripMenuItem";
            musicToolStripMenuItem.Size = new System.Drawing.Size(379, 44);
            musicToolStripMenuItem.Text = "&Music";
            // 
            // updatePlayerSpeedsToolStripMenuItem
            // 
            updatePlayerSpeedsToolStripMenuItem.Name = "updatePlayerSpeedsToolStripMenuItem";
            updatePlayerSpeedsToolStripMenuItem.Size = new System.Drawing.Size(379, 44);
            updatePlayerSpeedsToolStripMenuItem.Text = "&Update Player Speeds";
            // 
            // allowMultilinePasteToolStripMenuItem
            // 
            allowMultilinePasteToolStripMenuItem.Name = "allowMultilinePasteToolStripMenuItem";
            allowMultilinePasteToolStripMenuItem.Size = new System.Drawing.Size(379, 44);
            allowMultilinePasteToolStripMenuItem.Text = "Allow Multiline &Paste";
            // 
            // showVersionIconsToolStripMenuItem
            // 
            showVersionIconsToolStripMenuItem.Name = "showVersionIconsToolStripMenuItem";
            showVersionIconsToolStripMenuItem.Size = new System.Drawing.Size(379, 44);
            showVersionIconsToolStripMenuItem.Text = "Show &Version Icons";
            // 
            // colorNamesInListToolStripMenuItem
            // 
            colorNamesInListToolStripMenuItem.Name = "colorNamesInListToolStripMenuItem";
            colorNamesInListToolStripMenuItem.Size = new System.Drawing.Size(379, 44);
            colorNamesInListToolStripMenuItem.Text = "&Color Names in List";
            // 
            // toolStripMenuItem4
            // 
            toolStripMenuItem4.Name = "toolStripMenuItem4";
            toolStripMenuItem4.Size = new System.Drawing.Size(376, 6);
            // 
            // awayToolStripMenuItem
            // 
            awayToolStripMenuItem.Name = "awayToolStripMenuItem";
            awayToolStripMenuItem.Size = new System.Drawing.Size(379, 44);
            awayToolStripMenuItem.Text = "&Away";
            awayToolStripMenuItem.Click += awayToolStripMenuItem_Click;
            // 
            // windowToolStripMenuItem
            // 
            windowToolStripMenuItem.Name = "windowToolStripMenuItem";
            windowToolStripMenuItem.Size = new System.Drawing.Size(121, 36);
            windowToolStripMenuItem.Text = "&Window";
            // 
            // txtChat
            // 
            txtChat.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)0));
            txtChat.Location = new System.Drawing.Point(266, 60);
            txtChat.Margin = new System.Windows.Forms.Padding(6);
            txtChat.Name = "txtChat";
            txtChat.Size = new System.Drawing.Size(1422, 676);
            txtChat.TabIndex = 5;
            txtChat.Text = "";
            // 
            // ServerWindow
            // 
            AcceptButton = btnSend;
            AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            ClientSize = new System.Drawing.Size(1696, 795);
            Controls.Add(txtChat);
            Controls.Add(btnSend);
            Controls.Add(txtInput);
            Controls.Add(lstPlayers);
            Controls.Add(menuStrip1);
            Icon = ((System.Drawing.Icon)resources.GetObject("$this.Icon"));
            ImeMode = System.Windows.Forms.ImeMode.On;
            MainMenuStrip = menuStrip1;
            Margin = new System.Windows.Forms.Padding(6);
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Stadium:";
            FormClosing += ServerWindow_FormClosing;
            Load += ServerWindow_Load;
            playerMenu.ResumeLayout(false);
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        private System.Windows.Forms.RichTextBox txtChat;

        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem4;
        private System.Windows.Forms.ToolStripMenuItem soundToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem musicToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem updatePlayerSpeedsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem allowMultilinePasteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem showVersionIconsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem colorNamesInListToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem windowToolStripMenuItem;

        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem challengeInfoToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem watchBattleToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem privateMessageToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem controlWindowToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem kickToolStripMenuItem1;

        private System.Windows.Forms.ToolStripMenuItem playerToolStripMenuItem;

        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveLogToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem loadAnotherTeamToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem changeItemsToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem accessBoxToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem rearrangeTeamToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem teamToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openTeamBuilderToolStripMenuItem;

        #endregion

        private System.Windows.Forms.ListView lstPlayers;
        private System.Windows.Forms.TextBox txtInput;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.ColumnHeader nameColumn;
        private System.Windows.Forms.ContextMenuStrip playerMenu;
        private System.Windows.Forms.ToolStripMenuItem challengeInfoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem watchBattleToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem privateMessageToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem controlWindowToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem kickToolStripMenuItem;
        private System.Windows.Forms.MenuStrip menuStrip1;
        public System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
        public System.Windows.Forms.ToolStripMenuItem awayToolStripMenuItem;
        private System.Windows.Forms.ImageList imageList1;
    }
}