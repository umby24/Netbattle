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
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            ListViewItem listViewItem1 = new ListViewItem("umby24", 3);
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ServerWindow));
            lstPlayers = new ListView();
            nameColumn = new ColumnHeader();
            playerMenu = new ContextMenuStrip(components);
            challengeInfoToolStripMenuItem = new ToolStripMenuItem();
            watchBattleToolStripMenuItem = new ToolStripMenuItem();
            privateMessageToolStripMenuItem = new ToolStripMenuItem();
            controlWindowToolStripMenuItem = new ToolStripMenuItem();
            toolStripMenuItem1 = new ToolStripSeparator();
            kickToolStripMenuItem = new ToolStripMenuItem();
            imageList1 = new ImageList(components);
            txtInput = new TextBox();
            btnSend = new Button();
            menuStrip1 = new MenuStrip();
            fileToolStripMenuItem = new ToolStripMenuItem();
            saveLogToolStripMenuItem = new ToolStripMenuItem();
            toolStripMenuItem2 = new ToolStripSeparator();
            exitToolStripMenuItem = new ToolStripMenuItem();
            teamToolStripMenuItem = new ToolStripMenuItem();
            rearrangeTeamToolStripMenuItem = new ToolStripMenuItem();
            accessBoxToolStripMenuItem = new ToolStripMenuItem();
            changeItemsToolStripMenuItem = new ToolStripMenuItem();
            openTeamBuilderToolStripMenuItem = new ToolStripMenuItem();
            loadAnotherTeamToolStripMenuItem = new ToolStripMenuItem();
            playerToolStripMenuItem = new ToolStripMenuItem();
            challengeInfoToolStripMenuItem1 = new ToolStripMenuItem();
            watchBattleToolStripMenuItem1 = new ToolStripMenuItem();
            privateMessageToolStripMenuItem1 = new ToolStripMenuItem();
            controlWindowToolStripMenuItem1 = new ToolStripMenuItem();
            toolStripMenuItem3 = new ToolStripSeparator();
            kickToolStripMenuItem1 = new ToolStripMenuItem();
            optionsToolStripMenuItem = new ToolStripMenuItem();
            soundToolStripMenuItem = new ToolStripMenuItem();
            musicToolStripMenuItem = new ToolStripMenuItem();
            updatePlayerSpeedsToolStripMenuItem = new ToolStripMenuItem();
            allowMultilinePasteToolStripMenuItem = new ToolStripMenuItem();
            showVersionIconsToolStripMenuItem = new ToolStripMenuItem();
            colorNamesInListToolStripMenuItem = new ToolStripMenuItem();
            toolStripMenuItem4 = new ToolStripSeparator();
            awayToolStripMenuItem = new ToolStripMenuItem();
            windowToolStripMenuItem = new ToolStripMenuItem();
            txtChat = new RichTextBox();
            statusStrip1 = new StatusStrip();
            lblStatus = new ToolStripStatusLabel();
            lblBytes = new ToolStripStatusLabel();
            playerMenu.SuspendLayout();
            menuStrip1.SuspendLayout();
            statusStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // lstPlayers
            // 
            lstPlayers.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            lstPlayers.Columns.AddRange(new ColumnHeader[] { nameColumn });
            lstPlayers.ContextMenuStrip = playerMenu;
            lstPlayers.Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            lstPlayers.FullRowSelect = true;
            lstPlayers.HeaderStyle = ColumnHeaderStyle.None;
            lstPlayers.Items.AddRange(new ListViewItem[] { listViewItem1 });
            lstPlayers.LargeImageList = imageList1;
            lstPlayers.Location = new Point(0, 31);
            lstPlayers.Margin = new Padding(4);
            lstPlayers.MultiSelect = false;
            lstPlayers.Name = "lstPlayers";
            lstPlayers.Size = new Size(120, 350);
            lstPlayers.SmallImageList = imageList1;
            lstPlayers.TabIndex = 1;
            lstPlayers.TileSize = new Size(10, 50);
            lstPlayers.UseCompatibleStateImageBehavior = false;
            lstPlayers.View = View.List;
            // 
            // nameColumn
            // 
            nameColumn.Name = "nameColumn";
            nameColumn.Text = "Name";
            nameColumn.Width = 115;
            // 
            // playerMenu
            // 
            playerMenu.ImageScalingSize = new Size(32, 32);
            playerMenu.Items.AddRange(new ToolStripItem[] { challengeInfoToolStripMenuItem, watchBattleToolStripMenuItem, privateMessageToolStripMenuItem, controlWindowToolStripMenuItem, toolStripMenuItem1, kickToolStripMenuItem });
            playerMenu.Name = "playerMenu";
            playerMenu.Size = new Size(162, 120);
            // 
            // challengeInfoToolStripMenuItem
            // 
            challengeInfoToolStripMenuItem.Name = "challengeInfoToolStripMenuItem";
            challengeInfoToolStripMenuItem.Size = new Size(161, 22);
            challengeInfoToolStripMenuItem.Text = "Challenge/Info";
            challengeInfoToolStripMenuItem.Click += challengeInfoToolStripMenuItem_Click;
            // 
            // watchBattleToolStripMenuItem
            // 
            watchBattleToolStripMenuItem.Name = "watchBattleToolStripMenuItem";
            watchBattleToolStripMenuItem.Size = new Size(161, 22);
            watchBattleToolStripMenuItem.Text = "Watch Battle";
            watchBattleToolStripMenuItem.Click += watchBattleToolStripMenuItem_Click;
            // 
            // privateMessageToolStripMenuItem
            // 
            privateMessageToolStripMenuItem.Name = "privateMessageToolStripMenuItem";
            privateMessageToolStripMenuItem.Size = new Size(161, 22);
            privateMessageToolStripMenuItem.Text = "Private Message";
            privateMessageToolStripMenuItem.Click += privateMessageToolStripMenuItem_Click;
            // 
            // controlWindowToolStripMenuItem
            // 
            controlWindowToolStripMenuItem.Name = "controlWindowToolStripMenuItem";
            controlWindowToolStripMenuItem.Size = new Size(161, 22);
            controlWindowToolStripMenuItem.Text = "Control Window";
            controlWindowToolStripMenuItem.Click += controlWindowToolStripMenuItem_Click;
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new Size(158, 6);
            // 
            // kickToolStripMenuItem
            // 
            kickToolStripMenuItem.Name = "kickToolStripMenuItem";
            kickToolStripMenuItem.Size = new Size(161, 22);
            kickToolStripMenuItem.Text = "&Kick";
            kickToolStripMenuItem.Click += kickToolStripMenuItem_Click;
            // 
            // imageList1
            // 
            imageList1.ColorDepth = ColorDepth.Depth8Bit;
            imageList1.ImageStream = (ImageListStreamer)resources.GetObject("imageList1.ImageStream");
            imageList1.TransparentColor = Color.Transparent;
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
            txtInput.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            txtInput.Location = new Point(130, 352);
            txtInput.Margin = new Padding(6);
            txtInput.MaxLength = 200;
            txtInput.Name = "txtInput";
            txtInput.Size = new Size(325, 23);
            txtInput.TabIndex = 2;
            // 
            // btnSend
            // 
            btnSend.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnSend.Location = new Point(465, 351);
            btnSend.Margin = new Padding(6);
            btnSend.Name = "btnSend";
            btnSend.Size = new Size(73, 23);
            btnSend.TabIndex = 3;
            btnSend.Text = "Send";
            btnSend.UseVisualStyleBackColor = true;
            btnSend.Click += btnSend_Click;
            // 
            // menuStrip1
            // 
            menuStrip1.ImageScalingSize = new Size(32, 32);
            menuStrip1.Items.AddRange(new ToolStripItem[] { fileToolStripMenuItem, teamToolStripMenuItem, playerToolStripMenuItem, optionsToolStripMenuItem, windowToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Padding = new Padding(12, 4, 0, 4);
            menuStrip1.Size = new Size(553, 27);
            menuStrip1.TabIndex = 4;
            menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { saveLogToolStripMenuItem, toolStripMenuItem2, exitToolStripMenuItem });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new Size(37, 19);
            fileToolStripMenuItem.Text = "&File";
            // 
            // saveLogToolStripMenuItem
            // 
            saveLogToolStripMenuItem.Name = "saveLogToolStripMenuItem";
            saveLogToolStripMenuItem.Size = new Size(130, 22);
            saveLogToolStripMenuItem.Text = "&Save Log...";
            // 
            // toolStripMenuItem2
            // 
            toolStripMenuItem2.Name = "toolStripMenuItem2";
            toolStripMenuItem2.Size = new Size(127, 6);
            // 
            // exitToolStripMenuItem
            // 
            exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            exitToolStripMenuItem.Size = new Size(130, 22);
            exitToolStripMenuItem.Text = "E&xit";
            // 
            // teamToolStripMenuItem
            // 
            teamToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { rearrangeTeamToolStripMenuItem, accessBoxToolStripMenuItem, changeItemsToolStripMenuItem, openTeamBuilderToolStripMenuItem, loadAnotherTeamToolStripMenuItem });
            teamToolStripMenuItem.Name = "teamToolStripMenuItem";
            teamToolStripMenuItem.Size = new Size(47, 19);
            teamToolStripMenuItem.Text = "&Team";
            // 
            // rearrangeTeamToolStripMenuItem
            // 
            rearrangeTeamToolStripMenuItem.Name = "rearrangeTeamToolStripMenuItem";
            rearrangeTeamToolStripMenuItem.Size = new Size(177, 22);
            rearrangeTeamToolStripMenuItem.Text = "Rearrange Team";
            // 
            // accessBoxToolStripMenuItem
            // 
            accessBoxToolStripMenuItem.Name = "accessBoxToolStripMenuItem";
            accessBoxToolStripMenuItem.Size = new Size(177, 22);
            accessBoxToolStripMenuItem.Text = "Access Box";
            // 
            // changeItemsToolStripMenuItem
            // 
            changeItemsToolStripMenuItem.Name = "changeItemsToolStripMenuItem";
            changeItemsToolStripMenuItem.Size = new Size(177, 22);
            changeItemsToolStripMenuItem.Text = "Change Items";
            // 
            // openTeamBuilderToolStripMenuItem
            // 
            openTeamBuilderToolStripMenuItem.Name = "openTeamBuilderToolStripMenuItem";
            openTeamBuilderToolStripMenuItem.Size = new Size(177, 22);
            openTeamBuilderToolStripMenuItem.Text = "&Open Team Builder";
            openTeamBuilderToolStripMenuItem.Click += openTeamBuilderToolStripMenuItem_Click;
            // 
            // loadAnotherTeamToolStripMenuItem
            // 
            loadAnotherTeamToolStripMenuItem.Name = "loadAnotherTeamToolStripMenuItem";
            loadAnotherTeamToolStripMenuItem.Size = new Size(177, 22);
            loadAnotherTeamToolStripMenuItem.Text = "Load Another Team";
            // 
            // playerToolStripMenuItem
            // 
            playerToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { challengeInfoToolStripMenuItem1, watchBattleToolStripMenuItem1, privateMessageToolStripMenuItem1, controlWindowToolStripMenuItem1, toolStripMenuItem3, kickToolStripMenuItem1 });
            playerToolStripMenuItem.Name = "playerToolStripMenuItem";
            playerToolStripMenuItem.Size = new Size(51, 19);
            playerToolStripMenuItem.Text = "&Player";
            // 
            // challengeInfoToolStripMenuItem1
            // 
            challengeInfoToolStripMenuItem1.Name = "challengeInfoToolStripMenuItem1";
            challengeInfoToolStripMenuItem1.Size = new Size(168, 22);
            challengeInfoToolStripMenuItem1.Text = "&Challenge/Info...";
            // 
            // watchBattleToolStripMenuItem1
            // 
            watchBattleToolStripMenuItem1.Name = "watchBattleToolStripMenuItem1";
            watchBattleToolStripMenuItem1.Size = new Size(168, 22);
            watchBattleToolStripMenuItem1.Text = "&Watch Battle...";
            // 
            // privateMessageToolStripMenuItem1
            // 
            privateMessageToolStripMenuItem1.Name = "privateMessageToolStripMenuItem1";
            privateMessageToolStripMenuItem1.Size = new Size(168, 22);
            privateMessageToolStripMenuItem1.Text = "&Private Message...";
            // 
            // controlWindowToolStripMenuItem1
            // 
            controlWindowToolStripMenuItem1.Name = "controlWindowToolStripMenuItem1";
            controlWindowToolStripMenuItem1.Size = new Size(168, 22);
            controlWindowToolStripMenuItem1.Text = "C&ontrol Window";
            // 
            // toolStripMenuItem3
            // 
            toolStripMenuItem3.Name = "toolStripMenuItem3";
            toolStripMenuItem3.Size = new Size(165, 6);
            // 
            // kickToolStripMenuItem1
            // 
            kickToolStripMenuItem1.Name = "kickToolStripMenuItem1";
            kickToolStripMenuItem1.Size = new Size(168, 22);
            kickToolStripMenuItem1.Text = "&Kick";
            // 
            // optionsToolStripMenuItem
            // 
            optionsToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { soundToolStripMenuItem, musicToolStripMenuItem, updatePlayerSpeedsToolStripMenuItem, allowMultilinePasteToolStripMenuItem, showVersionIconsToolStripMenuItem, colorNamesInListToolStripMenuItem, toolStripMenuItem4, awayToolStripMenuItem });
            optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            optionsToolStripMenuItem.Size = new Size(61, 19);
            optionsToolStripMenuItem.Text = "&Options";
            // 
            // soundToolStripMenuItem
            // 
            soundToolStripMenuItem.Name = "soundToolStripMenuItem";
            soundToolStripMenuItem.Size = new Size(187, 22);
            soundToolStripMenuItem.Text = "&Sound";
            // 
            // musicToolStripMenuItem
            // 
            musicToolStripMenuItem.Name = "musicToolStripMenuItem";
            musicToolStripMenuItem.Size = new Size(187, 22);
            musicToolStripMenuItem.Text = "&Music";
            // 
            // updatePlayerSpeedsToolStripMenuItem
            // 
            updatePlayerSpeedsToolStripMenuItem.Name = "updatePlayerSpeedsToolStripMenuItem";
            updatePlayerSpeedsToolStripMenuItem.Size = new Size(187, 22);
            updatePlayerSpeedsToolStripMenuItem.Text = "&Update Player Speeds";
            // 
            // allowMultilinePasteToolStripMenuItem
            // 
            allowMultilinePasteToolStripMenuItem.Name = "allowMultilinePasteToolStripMenuItem";
            allowMultilinePasteToolStripMenuItem.Size = new Size(187, 22);
            allowMultilinePasteToolStripMenuItem.Text = "Allow Multiline &Paste";
            // 
            // showVersionIconsToolStripMenuItem
            // 
            showVersionIconsToolStripMenuItem.Name = "showVersionIconsToolStripMenuItem";
            showVersionIconsToolStripMenuItem.Size = new Size(187, 22);
            showVersionIconsToolStripMenuItem.Text = "Show &Version Icons";
            // 
            // colorNamesInListToolStripMenuItem
            // 
            colorNamesInListToolStripMenuItem.Name = "colorNamesInListToolStripMenuItem";
            colorNamesInListToolStripMenuItem.Size = new Size(187, 22);
            colorNamesInListToolStripMenuItem.Text = "&Color Names in List";
            // 
            // toolStripMenuItem4
            // 
            toolStripMenuItem4.Name = "toolStripMenuItem4";
            toolStripMenuItem4.Size = new Size(184, 6);
            // 
            // awayToolStripMenuItem
            // 
            awayToolStripMenuItem.Name = "awayToolStripMenuItem";
            awayToolStripMenuItem.Size = new Size(187, 22);
            awayToolStripMenuItem.Text = "&Away";
            awayToolStripMenuItem.Click += awayToolStripMenuItem_Click;
            // 
            // windowToolStripMenuItem
            // 
            windowToolStripMenuItem.Name = "windowToolStripMenuItem";
            windowToolStripMenuItem.Size = new Size(63, 19);
            windowToolStripMenuItem.Text = "&Window";
            // 
            // txtChat
            // 
            txtChat.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtChat.Font = new Font("Microsoft Sans Serif", 7.875F, FontStyle.Regular, GraphicsUnit.Point, 0);
            txtChat.Location = new Point(130, 31);
            txtChat.Margin = new Padding(6);
            txtChat.Name = "txtChat";
            txtChat.Size = new Size(410, 309);
            txtChat.TabIndex = 5;
            txtChat.Text = "";
            // 
            // statusStrip1
            // 
            statusStrip1.Items.AddRange(new ToolStripItem[] { lblStatus, lblBytes });
            statusStrip1.Location = new Point(0, 385);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new Size(553, 22);
            statusStrip1.TabIndex = 6;
            statusStrip1.Text = "statusStrip1";
            // 
            // lblStatus
            // 
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(118, 17);
            lblStatus.Text = "toolStripStatusLabel1";
            // 
            // lblBytes
            // 
            lblBytes.Name = "lblBytes";
            lblBytes.Size = new Size(118, 17);
            lblBytes.Text = "toolStripStatusLabel3";
            // 
            // ServerWindow
            // 
            AcceptButton = btnSend;
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(553, 407);
            Controls.Add(statusStrip1);
            Controls.Add(txtChat);
            Controls.Add(btnSend);
            Controls.Add(txtInput);
            Controls.Add(lstPlayers);
            Controls.Add(menuStrip1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            ImeMode = ImeMode.On;
            MainMenuStrip = menuStrip1;
            Margin = new Padding(6);
            Name = "ServerWindow";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Stadium:";
            FormClosing += ServerWindow_FormClosing;
            Load += ServerWindow_Load;
            playerMenu.ResumeLayout(false);
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        private System.Windows.Forms.ToolStripStatusLabel lblBytes;

        private System.Windows.Forms.ToolStripStatusLabel lblStatus;

        private System.Windows.Forms.StatusStrip statusStrip1;

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