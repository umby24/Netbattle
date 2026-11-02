namespace Netbattle.Forms {
    partial class TeamBuilder {
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
            ListViewItem listViewItem1 = new ListViewItem("#01", 0);
            ListViewItem listViewItem2 = new ListViewItem("#02", 1);
            ListViewItem listViewItem3 = new ListViewItem("#03", 2);
            ListViewItem listViewItem4 = new ListViewItem("#04", 3);
            ListViewItem listViewItem5 = new ListViewItem("#05", 4);
            ListViewItem listViewItem6 = new ListViewItem("#06", 5);
            ListViewItem listViewItem7 = new ListViewItem("#07", 6);
            ListViewItem listViewItem8 = new ListViewItem("#08", 7);
            ListViewItem listViewItem9 = new ListViewItem("#09", 8);
            ListViewItem listViewItem10 = new ListViewItem("#10", 9);
            ListViewItem listViewItem11 = new ListViewItem("#11", 10);
            ListViewItem listViewItem12 = new ListViewItem("#12", 11);
            ListViewItem listViewItem13 = new ListViewItem("#13", 12);
            ListViewItem listViewItem14 = new ListViewItem("#14", 13);
            ListViewItem listViewItem15 = new ListViewItem("#15", 14);
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TeamBuilder));
            tabControl1 = new TabControl();
            tbTrainer = new TabPage();
            grpAutoMessages = new GroupBox();
            txtLoseMessage = new TextBox();
            txtWinMessage = new TextBox();
            lblAutoLose = new Label();
            lblAutoWin = new Label();
            grpInfo = new GroupBox();
            dropGraphics = new ComboBox();
            txtExtraInfo = new TextBox();
            txtUsername = new TextBox();
            lblExtraInfo = new Label();
            lblGraphics = new Label();
            lblUsername = new Label();
            grpImage = new GroupBox();
            listView1 = new ListView();
            imageList2 = new ImageList(components);
            tbSlot1 = new TabPage();
            btnPoke6 = new Button();
            btnPoke5 = new Button();
            btnPoke4 = new Button();
            btnPoke3 = new Button();
            btnPoke2 = new Button();
            btnPoke1 = new Button();
            dropItem = new ComboBox();
            txtNickname = new TextBox();
            txtMove4 = new TextBox();
            txtMove3 = new TextBox();
            txtMove2 = new TextBox();
            txtMove1 = new TextBox();
            lblLevel = new Label();
            grpStats = new GroupBox();
            lblSpecialDefense = new Label();
            lblSpecialAttack = new Label();
            lblSpeed = new Label();
            lblDefense = new Label();
            lblAttack = new Label();
            lblHp = new Label();
            lblTypes = new Label();
            listMoves = new ListView();
            colMove = new ColumnHeader();
            colPower = new ColumnHeader();
            colAcc = new ColumnHeader();
            colPP = new ColumnHeader();
            colLearned = new ColumnHeader();
            imageList1 = new ImageList(components);
            label2 = new Label();
            lblNickname = new Label();
            dropPokemon = new ComboBox();
            btnSwitch = new Button();
            lblPokemon = new Label();
            picImage = new PictureBox();
            menuStrip1 = new MenuStrip();
            fileToolStripMenuItem = new ToolStripMenuItem();
            openToolStripMenuItem = new ToolStripMenuItem();
            saveToolStripMenuItem = new ToolStripMenuItem();
            versionToolStripMenuItem = new ToolStripMenuItem();
            trueRBYToolStripMenuItem = new ToolStripMenuItem();
            rBYWithTradesToolStripMenuItem = new ToolStripMenuItem();
            toolStripMenuItem1 = new ToolStripSeparator();
            trueGSCToolStripMenuItem = new ToolStripMenuItem();
            gSCWithTradesToolStripMenuItem = new ToolStripMenuItem();
            toolStripMenuItem2 = new ToolStripSeparator();
            ruSaOnlyToolStripMenuItem = new ToolStripMenuItem();
            fullAdvancedToolStripMenuItem = new ToolStripMenuItem();
            toolStripMenuItem3 = new ToolStripSeparator();
            noModLoadedToolStripMenuItem = new ToolStripMenuItem();
            viewCurrentModsToolStripMenuItem = new ToolStripMenuItem();
            loadNewModFileToolStripMenuItem = new ToolStripMenuItem();
            sortToolStripMenuItem = new ToolStripMenuItem();
            dataDexToolStripMenuItem = new ToolStripMenuItem();
            pokedexToolStripMenuItem = new ToolStripMenuItem();
            moveDexToolStripMenuItem = new ToolStripMenuItem();
            battleDexToolStripMenuItem = new ToolStripMenuItem();
            damageCalcToolStripMenuItem = new ToolStripMenuItem();
            ofdPnb = new OpenFileDialog();
            tabControl1.SuspendLayout();
            tbTrainer.SuspendLayout();
            grpAutoMessages.SuspendLayout();
            grpInfo.SuspendLayout();
            grpImage.SuspendLayout();
            tbSlot1.SuspendLayout();
            grpStats.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)picImage).BeginInit();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(tbTrainer);
            tabControl1.Controls.Add(tbSlot1);
            tabControl1.Dock = DockStyle.Fill;
            tabControl1.Location = new Point(0, 24);
            tabControl1.Margin = new Padding(6, 7, 6, 7);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(730, 621);
            tabControl1.TabIndex = 0;
            // 
            // tbTrainer
            // 
            tbTrainer.Controls.Add(grpAutoMessages);
            tbTrainer.Controls.Add(grpInfo);
            tbTrainer.Controls.Add(grpImage);
            tbTrainer.Location = new Point(4, 24);
            tbTrainer.Margin = new Padding(6, 7, 6, 7);
            tbTrainer.Name = "tbTrainer";
            tbTrainer.Padding = new Padding(6, 7, 6, 7);
            tbTrainer.Size = new Size(722, 593);
            tbTrainer.TabIndex = 0;
            tbTrainer.Text = "Trainer Info";
            tbTrainer.UseVisualStyleBackColor = true;
            // 
            // grpAutoMessages
            // 
            grpAutoMessages.Controls.Add(txtLoseMessage);
            grpAutoMessages.Controls.Add(txtWinMessage);
            grpAutoMessages.Controls.Add(lblAutoLose);
            grpAutoMessages.Controls.Add(lblAutoWin);
            grpAutoMessages.Location = new Point(185, 272);
            grpAutoMessages.Margin = new Padding(6, 7, 6, 7);
            grpAutoMessages.Name = "grpAutoMessages";
            grpAutoMessages.Padding = new Padding(6, 7, 6, 7);
            grpAutoMessages.Size = new Size(525, 308);
            grpAutoMessages.TabIndex = 3;
            grpAutoMessages.TabStop = false;
            grpAutoMessages.Text = "Auto Messages";
            // 
            // txtLoseMessage
            // 
            txtLoseMessage.Location = new Point(12, 179);
            txtLoseMessage.Margin = new Padding(6, 7, 6, 7);
            txtLoseMessage.MaxLength = 200;
            txtLoseMessage.Multiline = true;
            txtLoseMessage.Name = "txtLoseMessage";
            txtLoseMessage.ScrollBars = ScrollBars.Vertical;
            txtLoseMessage.Size = new Size(503, 105);
            txtLoseMessage.TabIndex = 6;
            txtLoseMessage.Text = "wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww";
            // 
            // txtWinMessage
            // 
            txtWinMessage.Location = new Point(10, 45);
            txtWinMessage.Margin = new Padding(6, 7, 6, 7);
            txtWinMessage.MaxLength = 200;
            txtWinMessage.Multiline = true;
            txtWinMessage.Name = "txtWinMessage";
            txtWinMessage.ScrollBars = ScrollBars.Vertical;
            txtWinMessage.Size = new Size(505, 105);
            txtWinMessage.TabIndex = 5;
            txtWinMessage.Text = "wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww";
            // 
            // lblAutoLose
            // 
            lblAutoLose.AutoSize = true;
            lblAutoLose.Location = new Point(10, 157);
            lblAutoLose.Margin = new Padding(6, 0, 6, 0);
            lblAutoLose.Name = "lblAutoLose";
            lblAutoLose.Size = new Size(31, 15);
            lblAutoLose.TabIndex = 1;
            lblAutoLose.Text = "Lose";
            // 
            // lblAutoWin
            // 
            lblAutoWin.AutoSize = true;
            lblAutoWin.Location = new Point(12, 23);
            lblAutoWin.Margin = new Padding(6, 0, 6, 0);
            lblAutoWin.Name = "lblAutoWin";
            lblAutoWin.Size = new Size(28, 15);
            lblAutoWin.TabIndex = 0;
            lblAutoWin.Text = "Win";
            // 
            // grpInfo
            // 
            grpInfo.Controls.Add(dropGraphics);
            grpInfo.Controls.Add(txtExtraInfo);
            grpInfo.Controls.Add(txtUsername);
            grpInfo.Controls.Add(lblExtraInfo);
            grpInfo.Controls.Add(lblGraphics);
            grpInfo.Controls.Add(lblUsername);
            grpInfo.Location = new Point(186, 15);
            grpInfo.Margin = new Padding(6, 7, 6, 7);
            grpInfo.Name = "grpInfo";
            grpInfo.Padding = new Padding(6, 7, 6, 7);
            grpInfo.Size = new Size(524, 243);
            grpInfo.TabIndex = 2;
            grpInfo.TabStop = false;
            grpInfo.Text = "Info";
            // 
            // dropGraphics
            // 
            dropGraphics.FormattingEnabled = true;
            dropGraphics.Items.AddRange(new object[] { "Green", "Red/Blue", "Yellow", "Gold", "Silver", "Ruby/Sapphire", "Leaf/Fire", "Emerald" });
            dropGraphics.Location = new Point(162, 45);
            dropGraphics.Margin = new Padding(6, 7, 6, 7);
            dropGraphics.Name = "dropGraphics";
            dropGraphics.Size = new Size(97, 23);
            dropGraphics.TabIndex = 5;
            dropGraphics.SelectedIndexChanged += dropGraphics_SelectedIndexChanged;
            // 
            // txtExtraInfo
            // 
            txtExtraInfo.Location = new Point(12, 97);
            txtExtraInfo.Margin = new Padding(6, 7, 6, 7);
            txtExtraInfo.MaxLength = 200;
            txtExtraInfo.Multiline = true;
            txtExtraInfo.Name = "txtExtraInfo";
            txtExtraInfo.ScrollBars = ScrollBars.Vertical;
            txtExtraInfo.Size = new Size(502, 120);
            txtExtraInfo.TabIndex = 4;
            txtExtraInfo.Text = "wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww";
            // 
            // txtUsername
            // 
            txtUsername.Location = new Point(12, 45);
            txtUsername.Margin = new Padding(6, 7, 6, 7);
            txtUsername.MaxLength = 20;
            txtUsername.Name = "txtUsername";
            txtUsername.Size = new Size(138, 23);
            txtUsername.TabIndex = 3;
            // 
            // lblExtraInfo
            // 
            lblExtraInfo.AutoSize = true;
            lblExtraInfo.Location = new Point(12, 75);
            lblExtraInfo.Margin = new Padding(6, 0, 6, 0);
            lblExtraInfo.Name = "lblExtraInfo";
            lblExtraInfo.Size = new Size(57, 15);
            lblExtraInfo.TabIndex = 2;
            lblExtraInfo.Text = "Extra Info";
            // 
            // lblGraphics
            // 
            lblGraphics.AutoSize = true;
            lblGraphics.Location = new Point(162, 23);
            lblGraphics.Margin = new Padding(6, 0, 6, 0);
            lblGraphics.Name = "lblGraphics";
            lblGraphics.Size = new Size(53, 15);
            lblGraphics.TabIndex = 1;
            lblGraphics.Text = "Graphics";
            // 
            // lblUsername
            // 
            lblUsername.AutoSize = true;
            lblUsername.Location = new Point(12, 23);
            lblUsername.Margin = new Padding(6, 0, 6, 0);
            lblUsername.Name = "lblUsername";
            lblUsername.Size = new Size(60, 15);
            lblUsername.TabIndex = 0;
            lblUsername.Text = "Username";
            // 
            // grpImage
            // 
            grpImage.Controls.Add(listView1);
            grpImage.Location = new Point(11, 15);
            grpImage.Margin = new Padding(6, 7, 6, 7);
            grpImage.Name = "grpImage";
            grpImage.Padding = new Padding(6, 7, 6, 7);
            grpImage.Size = new Size(156, 566);
            grpImage.TabIndex = 1;
            grpImage.TabStop = false;
            grpImage.Text = "Image";
            // 
            // listView1
            // 
            listView1.Font = new Font("Microsoft Sans Serif", 10.125F, FontStyle.Regular, GraphicsUnit.Point, 0);
            listView1.FullRowSelect = true;
            listView1.ImeMode = ImeMode.NoControl;
            listView1.Items.AddRange(new ListViewItem[] { listViewItem1, listViewItem2, listViewItem3, listViewItem4, listViewItem5, listViewItem6, listViewItem7, listViewItem8, listViewItem9, listViewItem10, listViewItem11, listViewItem12, listViewItem13, listViewItem14, listViewItem15 });
            listView1.LargeImageList = imageList2;
            listView1.Location = new Point(13, 23);
            listView1.Margin = new Padding(6, 7, 6, 7);
            listView1.MultiSelect = false;
            listView1.Name = "listView1";
            listView1.Size = new Size(131, 420);
            listView1.SmallImageList = imageList2;
            listView1.TabIndex = 7;
            listView1.UseCompatibleStateImageBehavior = false;
            listView1.View = View.List;
            listView1.SelectedIndexChanged += listView1_SelectedIndexChanged;
            // 
            // imageList2
            // 
            imageList2.ColorDepth = ColorDepth.Depth8Bit;
            imageList2.ImageStream = (ImageListStreamer)resources.GetObject("imageList2.ImageStream");
            imageList2.TransparentColor = Color.Transparent;
            imageList2.Images.SetKeyName(0, "trainer1.ico");
            imageList2.Images.SetKeyName(1, "trainer2.ico");
            imageList2.Images.SetKeyName(2, "trainer3.ico");
            imageList2.Images.SetKeyName(3, "trainer4.ico");
            imageList2.Images.SetKeyName(4, "trainer5.ico");
            imageList2.Images.SetKeyName(5, "trainer6.ico");
            imageList2.Images.SetKeyName(6, "trainer7.ico");
            imageList2.Images.SetKeyName(7, "trainer8.ico");
            imageList2.Images.SetKeyName(8, "trainer9.ico");
            imageList2.Images.SetKeyName(9, "trainer10.ico");
            imageList2.Images.SetKeyName(10, "trainer11.ico");
            imageList2.Images.SetKeyName(11, "trainer12.ico");
            imageList2.Images.SetKeyName(12, "trainer13.ico");
            imageList2.Images.SetKeyName(13, "trainer14.ico");
            imageList2.Images.SetKeyName(14, "trainer15.ico");
            imageList2.Images.SetKeyName(15, "greyPokeball.ico");
            imageList2.Images.SetKeyName(16, "Icon22.ico");
            // 
            // tbSlot1
            // 
            tbSlot1.Controls.Add(btnPoke6);
            tbSlot1.Controls.Add(btnPoke5);
            tbSlot1.Controls.Add(btnPoke4);
            tbSlot1.Controls.Add(btnPoke3);
            tbSlot1.Controls.Add(btnPoke2);
            tbSlot1.Controls.Add(btnPoke1);
            tbSlot1.Controls.Add(dropItem);
            tbSlot1.Controls.Add(txtNickname);
            tbSlot1.Controls.Add(txtMove4);
            tbSlot1.Controls.Add(txtMove3);
            tbSlot1.Controls.Add(txtMove2);
            tbSlot1.Controls.Add(txtMove1);
            tbSlot1.Controls.Add(lblLevel);
            tbSlot1.Controls.Add(grpStats);
            tbSlot1.Controls.Add(listMoves);
            tbSlot1.Controls.Add(label2);
            tbSlot1.Controls.Add(lblNickname);
            tbSlot1.Controls.Add(dropPokemon);
            tbSlot1.Controls.Add(btnSwitch);
            tbSlot1.Controls.Add(lblPokemon);
            tbSlot1.Controls.Add(picImage);
            tbSlot1.ImageKey = "(none)";
            tbSlot1.Location = new Point(4, 24);
            tbSlot1.Margin = new Padding(6, 7, 6, 7);
            tbSlot1.Name = "tbSlot1";
            tbSlot1.Padding = new Padding(6, 7, 6, 7);
            tbSlot1.Size = new Size(722, 593);
            tbSlot1.TabIndex = 1;
            tbSlot1.Text = "Team";
            tbSlot1.UseVisualStyleBackColor = true;
            // 
            // btnPoke6
            // 
            btnPoke6.Font = new Font("Microsoft Sans Serif", 8.25F);
            btnPoke6.ImageAlign = ContentAlignment.MiddleLeft;
            btnPoke6.ImageIndex = 15;
            btnPoke6.ImageList = imageList2;
            btnPoke6.Location = new Point(482, 15);
            btnPoke6.Margin = new Padding(6, 7, 6, 7);
            btnPoke6.Name = "btnPoke6";
            btnPoke6.Size = new Size(82, 26);
            btnPoke6.TabIndex = 25;
            btnPoke6.Text = "PKMN 6";
            btnPoke6.TextAlign = ContentAlignment.MiddleRight;
            btnPoke6.UseVisualStyleBackColor = true;
            btnPoke6.Click += SwitchPokemonButtonClick;
            // 
            // btnPoke5
            // 
            btnPoke5.Font = new Font("Microsoft Sans Serif", 8.25F);
            btnPoke5.ImageAlign = ContentAlignment.MiddleLeft;
            btnPoke5.ImageIndex = 15;
            btnPoke5.ImageList = imageList2;
            btnPoke5.Location = new Point(388, 15);
            btnPoke5.Margin = new Padding(6, 7, 6, 7);
            btnPoke5.Name = "btnPoke5";
            btnPoke5.Size = new Size(82, 26);
            btnPoke5.TabIndex = 24;
            btnPoke5.Text = "PKMN 5";
            btnPoke5.TextAlign = ContentAlignment.MiddleRight;
            btnPoke5.UseVisualStyleBackColor = true;
            btnPoke5.Click += SwitchPokemonButtonClick;
            // 
            // btnPoke4
            // 
            btnPoke4.Font = new Font("Microsoft Sans Serif", 8.25F);
            btnPoke4.ImageAlign = ContentAlignment.MiddleLeft;
            btnPoke4.ImageIndex = 15;
            btnPoke4.ImageList = imageList2;
            btnPoke4.Location = new Point(294, 14);
            btnPoke4.Margin = new Padding(6, 7, 6, 7);
            btnPoke4.Name = "btnPoke4";
            btnPoke4.Size = new Size(82, 26);
            btnPoke4.TabIndex = 23;
            btnPoke4.Text = "PKMN 4";
            btnPoke4.TextAlign = ContentAlignment.MiddleRight;
            btnPoke4.UseVisualStyleBackColor = true;
            btnPoke4.Click += SwitchPokemonButtonClick;
            // 
            // btnPoke3
            // 
            btnPoke3.Font = new Font("Microsoft Sans Serif", 8.25F);
            btnPoke3.ImageAlign = ContentAlignment.MiddleLeft;
            btnPoke3.ImageIndex = 15;
            btnPoke3.ImageList = imageList2;
            btnPoke3.Location = new Point(200, 14);
            btnPoke3.Margin = new Padding(6, 7, 6, 7);
            btnPoke3.Name = "btnPoke3";
            btnPoke3.Size = new Size(82, 26);
            btnPoke3.TabIndex = 22;
            btnPoke3.Text = "PKMN 3";
            btnPoke3.TextAlign = ContentAlignment.MiddleRight;
            btnPoke3.UseVisualStyleBackColor = true;
            btnPoke3.Click += SwitchPokemonButtonClick;
            // 
            // btnPoke2
            // 
            btnPoke2.Font = new Font("Microsoft Sans Serif", 8.25F);
            btnPoke2.ImageAlign = ContentAlignment.MiddleLeft;
            btnPoke2.ImageIndex = 15;
            btnPoke2.ImageList = imageList2;
            btnPoke2.Location = new Point(106, 14);
            btnPoke2.Margin = new Padding(6, 7, 6, 7);
            btnPoke2.Name = "btnPoke2";
            btnPoke2.Size = new Size(82, 26);
            btnPoke2.TabIndex = 21;
            btnPoke2.Text = "PKMN 2";
            btnPoke2.TextAlign = ContentAlignment.MiddleRight;
            btnPoke2.UseVisualStyleBackColor = true;
            btnPoke2.Click += SwitchPokemonButtonClick;
            // 
            // btnPoke1
            // 
            btnPoke1.FlatStyle = FlatStyle.Popup;
            btnPoke1.Font = new Font("Microsoft Sans Serif", 8.25F);
            btnPoke1.ImageAlign = ContentAlignment.MiddleLeft;
            btnPoke1.ImageKey = "greyPokeball.ico";
            btnPoke1.ImageList = imageList2;
            btnPoke1.Location = new Point(12, 14);
            btnPoke1.Margin = new Padding(6, 7, 6, 7);
            btnPoke1.Name = "btnPoke1";
            btnPoke1.Size = new Size(82, 26);
            btnPoke1.TabIndex = 20;
            btnPoke1.Text = "PKMN 1";
            btnPoke1.TextAlign = ContentAlignment.MiddleRight;
            btnPoke1.UseVisualStyleBackColor = true;
            btnPoke1.Click += SwitchPokemonButtonClick;
            // 
            // dropItem
            // 
            dropItem.Font = new Font("Microsoft Sans Serif", 8.25F);
            dropItem.FormattingEnabled = true;
            dropItem.Items.AddRange(new object[] { "No Item", "Berry ", "Berry Juice", "Bitter Berry", "Burnt Berry", "Gold Berry", "Ice Berry", "Mint Berry", "Miracle Berry", "Mystery Berry", "Paralyzecure Berry", "Poisoncure Berry", "Berserk Gene", "Black Belt", "Black Glasses", "Bright Powder", "Charcoal ", "Dragon Fang", "Focus Band", "Hard Stone", "Kings Rock", "Leftovers ", "Light Ball", "Lucky Punch", "Magnet ", "Metal Coat", "Metal Powder", "Miracle Seed", "Mystic Water", "Nevermelt Ice", "Pink Bow", "Poison Barb", "Polkadot Bow", "Quick Claw", "Scope Lens", "Sharp Beak", "Silver Powder", "Soft Sand", "Spell Tag", "Stick", "Thick Club", "Twisted Spoon", "Cheri Berry", "Chesto Berry", "Pecha Berry", "Rawst Berry", "Aspear Berry", "Leppa Berry", "Oran Berry", "Persim Berry", "Lum Berry", "Sitrus Berry", "Figy Berry", "Iapapa Berry", "Mago Berry", "Wiki Berry", "Aguav Berry", "Liechi Berry", "Ganlon Berry", "Salac Berry", "Petaya Berry", "Apicot Berry", "Lansat Berry", "Starf Berry", "Choice Band", "Deep Sea Scale", "Deep Sea Tooth", "Lax Incense", "Macho Brace", "Mental Herb", "Sea Incense", "Shell Bell", "Silk Scarf", "Soul Dew", "White Herb" });
            dropItem.Location = new Point(12, 174);
            dropItem.Margin = new Padding(6, 7, 6, 7);
            dropItem.Name = "dropItem";
            dropItem.Size = new Size(208, 21);
            dropItem.TabIndex = 19;
            dropItem.SelectedIndexChanged += dropItem_SelectedIndexChanged;
            // 
            // txtNickname
            // 
            txtNickname.Font = new Font("Microsoft Sans Serif", 8.25F);
            txtNickname.Location = new Point(12, 122);
            txtNickname.Margin = new Padding(6, 7, 6, 7);
            txtNickname.MaxLength = 15;
            txtNickname.Name = "txtNickname";
            txtNickname.Size = new Size(212, 20);
            txtNickname.TabIndex = 18;
            txtNickname.TextChanged += txtNickname_TextChanged;
            // 
            // txtMove4
            // 
            txtMove4.Font = new Font("Microsoft Sans Serif", 7.875F);
            txtMove4.Location = new Point(476, 551);
            txtMove4.Margin = new Padding(6, 7, 6, 7);
            txtMove4.Name = "txtMove4";
            txtMove4.ReadOnly = true;
            txtMove4.Size = new Size(163, 19);
            txtMove4.TabIndex = 17;
            // 
            // txtMove3
            // 
            txtMove3.Font = new Font("Microsoft Sans Serif", 7.875F);
            txtMove3.Location = new Point(302, 551);
            txtMove3.Margin = new Padding(6, 7, 6, 7);
            txtMove3.Name = "txtMove3";
            txtMove3.ReadOnly = true;
            txtMove3.Size = new Size(163, 19);
            txtMove3.TabIndex = 16;
            // 
            // txtMove2
            // 
            txtMove2.Font = new Font("Microsoft Sans Serif", 7.875F);
            txtMove2.Location = new Point(159, 551);
            txtMove2.Margin = new Padding(6, 7, 6, 7);
            txtMove2.Name = "txtMove2";
            txtMove2.ReadOnly = true;
            txtMove2.Size = new Size(131, 19);
            txtMove2.TabIndex = 15;
            // 
            // txtMove1
            // 
            txtMove1.Font = new Font("Microsoft Sans Serif", 7.875F, FontStyle.Regular, GraphicsUnit.Point, 0);
            txtMove1.Location = new Point(12, 551);
            txtMove1.Margin = new Padding(6, 7, 6, 7);
            txtMove1.Name = "txtMove1";
            txtMove1.ReadOnly = true;
            txtMove1.Size = new Size(132, 19);
            txtMove1.TabIndex = 14;
            txtMove1.Text = "MoveText";
            // 
            // lblLevel
            // 
            lblLevel.AutoSize = true;
            lblLevel.Font = new Font("Microsoft Sans Serif", 8.25F);
            lblLevel.Location = new Point(492, 175);
            lblLevel.Margin = new Padding(6, 0, 6, 0);
            lblLevel.Name = "lblLevel";
            lblLevel.Size = new Size(72, 13);
            lblLevel.TabIndex = 13;
            lblLevel.Text = "Lv. 100, Male";
            // 
            // grpStats
            // 
            grpStats.Controls.Add(lblSpecialDefense);
            grpStats.Controls.Add(lblSpecialAttack);
            grpStats.Controls.Add(lblSpeed);
            grpStats.Controls.Add(lblDefense);
            grpStats.Controls.Add(lblAttack);
            grpStats.Controls.Add(lblHp);
            grpStats.Controls.Add(lblTypes);
            grpStats.Location = new Point(236, 55);
            grpStats.Margin = new Padding(6, 7, 6, 7);
            grpStats.Name = "grpStats";
            grpStats.Padding = new Padding(6, 7, 6, 7);
            grpStats.Size = new Size(215, 157);
            grpStats.TabIndex = 7;
            grpStats.TabStop = false;
            grpStats.Text = "Stats";
            // 
            // lblSpecialDefense
            // 
            lblSpecialDefense.AutoSize = true;
            lblSpecialDefense.Font = new Font("Microsoft Sans Serif", 8.25F);
            lblSpecialDefense.Location = new Point(12, 99);
            lblSpecialDefense.Margin = new Padding(6, 0, 6, 0);
            lblSpecialDefense.Name = "lblSpecialDefense";
            lblSpecialDefense.Size = new Size(66, 13);
            lblSpecialDefense.TabIndex = 11;
            lblSpecialDefense.Text = "Sp. Defense";
            // 
            // lblSpecialAttack
            // 
            lblSpecialAttack.AutoSize = true;
            lblSpecialAttack.Font = new Font("Microsoft Sans Serif", 8.25F);
            lblSpecialAttack.Location = new Point(12, 84);
            lblSpecialAttack.Margin = new Padding(6, 0, 6, 0);
            lblSpecialAttack.Name = "lblSpecialAttack";
            lblSpecialAttack.Size = new Size(57, 13);
            lblSpecialAttack.TabIndex = 10;
            lblSpecialAttack.Text = "Sp. Attack";
            // 
            // lblSpeed
            // 
            lblSpeed.AutoSize = true;
            lblSpeed.Font = new Font("Microsoft Sans Serif", 8.25F);
            lblSpeed.Location = new Point(12, 69);
            lblSpeed.Margin = new Padding(6, 0, 6, 0);
            lblSpeed.Name = "lblSpeed";
            lblSpeed.Size = new Size(38, 13);
            lblSpeed.TabIndex = 9;
            lblSpeed.Text = "Speed";
            // 
            // lblDefense
            // 
            lblDefense.AutoSize = true;
            lblDefense.Font = new Font("Microsoft Sans Serif", 8.25F);
            lblDefense.Location = new Point(12, 53);
            lblDefense.Margin = new Padding(6, 0, 6, 0);
            lblDefense.Name = "lblDefense";
            lblDefense.Size = new Size(47, 13);
            lblDefense.TabIndex = 8;
            lblDefense.Text = "Defense";
            // 
            // lblAttack
            // 
            lblAttack.AutoSize = true;
            lblAttack.Font = new Font("Microsoft Sans Serif", 8.25F);
            lblAttack.Location = new Point(12, 38);
            lblAttack.Margin = new Padding(6, 0, 6, 0);
            lblAttack.Name = "lblAttack";
            lblAttack.Size = new Size(38, 13);
            lblAttack.TabIndex = 7;
            lblAttack.Text = "Attack";
            // 
            // lblHp
            // 
            lblHp.AutoSize = true;
            lblHp.Font = new Font("Microsoft Sans Serif", 8.25F);
            lblHp.Location = new Point(12, 23);
            lblHp.Margin = new Padding(6, 0, 6, 0);
            lblHp.Name = "lblHp";
            lblHp.Size = new Size(22, 13);
            lblHp.TabIndex = 6;
            lblHp.Text = "HP";
            // 
            // lblTypes
            // 
            lblTypes.AutoSize = true;
            lblTypes.Font = new Font("Microsoft Sans Serif", 8.25F);
            lblTypes.Location = new Point(12, 114);
            lblTypes.Margin = new Padding(6, 0, 6, 0);
            lblTypes.Name = "lblTypes";
            lblTypes.Size = new Size(34, 13);
            lblTypes.TabIndex = 12;
            lblTypes.Text = "Type:";
            // 
            // listMoves
            // 
            listMoves.CheckBoxes = true;
            listMoves.Columns.AddRange(new ColumnHeader[] { colMove, colPower, colAcc, colPP, colLearned });
            listMoves.Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            listMoves.FullRowSelect = true;
            listMoves.LargeImageList = imageList1;
            listMoves.Location = new Point(12, 240);
            listMoves.Margin = new Padding(6, 7, 6, 7);
            listMoves.Name = "listMoves";
            listMoves.Size = new Size(627, 297);
            listMoves.SmallImageList = imageList1;
            listMoves.Sorting = SortOrder.Ascending;
            listMoves.TabIndex = 6;
            listMoves.UseCompatibleStateImageBehavior = false;
            listMoves.View = View.Details;
            listMoves.ItemChecked += listMoves_ItemChecked;
            // 
            // colMove
            // 
            colMove.Name = "colMove";
            colMove.Text = "Move";
            colMove.Width = 190;
            // 
            // colPower
            // 
            colPower.Name = "colPower";
            colPower.Text = "Power";
            colPower.Width = 120;
            // 
            // colAcc
            // 
            colAcc.Name = "colAcc";
            colAcc.Text = "Acc.";
            colAcc.Width = 83;
            // 
            // colPP
            // 
            colPP.Name = "colPP";
            colPP.Text = "PP";
            colPP.Width = 53;
            // 
            // colLearned
            // 
            colLearned.Name = "colLearned";
            colLearned.Text = "Learned By";
            colLearned.Width = 167;
            // 
            // imageList1
            // 
            imageList1.ColorDepth = ColorDepth.Depth8Bit;
            imageList1.ImageStream = (ImageListStreamer)resources.GetObject("imageList1.ImageStream");
            imageList1.TransparentColor = Color.Transparent;
            imageList1.Images.SetKeyName(0, "typeBug.ico");
            imageList1.Images.SetKeyName(1, "typeDark.ico");
            imageList1.Images.SetKeyName(2, "typeDragon.ico");
            imageList1.Images.SetKeyName(3, "typeElec.ico");
            imageList1.Images.SetKeyName(4, "typeFighting.ico");
            imageList1.Images.SetKeyName(5, "typeFire.ico");
            imageList1.Images.SetKeyName(6, "typeFlying.ico");
            imageList1.Images.SetKeyName(7, "typeGhost.ico");
            imageList1.Images.SetKeyName(8, "typeGrass.ico");
            imageList1.Images.SetKeyName(9, "typeGround.ico");
            imageList1.Images.SetKeyName(10, "typeIce.ico");
            imageList1.Images.SetKeyName(11, "typeNormal.ico");
            imageList1.Images.SetKeyName(12, "typePoison.ico");
            imageList1.Images.SetKeyName(13, "typePsychic.ico");
            imageList1.Images.SetKeyName(14, "typeRock.ico");
            imageList1.Images.SetKeyName(15, "typeSteel.ico");
            imageList1.Images.SetKeyName(16, "typeWater.ico");
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(12, 152);
            label2.Margin = new Padding(6, 0, 6, 0);
            label2.Name = "label2";
            label2.Size = new Size(31, 15);
            label2.TabIndex = 5;
            label2.Text = "Item";
            // 
            // lblNickname
            // 
            lblNickname.AutoSize = true;
            lblNickname.Location = new Point(9, 100);
            lblNickname.Margin = new Padding(6, 0, 6, 0);
            lblNickname.Name = "lblNickname";
            lblNickname.Size = new Size(61, 15);
            lblNickname.TabIndex = 4;
            lblNickname.Text = "Nickname";
            // 
            // dropPokemon
            // 
            dropPokemon.Font = new Font("Microsoft Sans Serif", 8.25F);
            dropPokemon.FormattingEnabled = true;
            dropPokemon.Location = new Point(12, 70);
            dropPokemon.Margin = new Padding(6, 7, 6, 7);
            dropPokemon.Name = "dropPokemon";
            dropPokemon.Size = new Size(124, 21);
            dropPokemon.TabIndex = 3;
            dropPokemon.SelectedIndexChanged += dropPokemon_SelectedIndexChanged;
            // 
            // btnSwitch
            // 
            btnSwitch.Location = new Point(148, 70);
            btnSwitch.Margin = new Padding(6, 7, 6, 7);
            btnSwitch.Name = "btnSwitch";
            btnSwitch.Size = new Size(59, 23);
            btnSwitch.TabIndex = 2;
            btnSwitch.Text = "Expert";
            btnSwitch.UseVisualStyleBackColor = true;
            btnSwitch.Click += btnSwitch_Click;
            // 
            // lblPokemon
            // 
            lblPokemon.AutoSize = true;
            lblPokemon.Location = new Point(9, 48);
            lblPokemon.Margin = new Padding(6, 0, 6, 0);
            lblPokemon.Name = "lblPokemon";
            lblPokemon.Size = new Size(58, 15);
            lblPokemon.TabIndex = 1;
            lblPokemon.Text = "Pokemon";
            // 
            // picImage
            // 
            picImage.Location = new Point(471, 69);
            picImage.Margin = new Padding(6, 7, 6, 7);
            picImage.Name = "picImage";
            picImage.Size = new Size(117, 99);
            picImage.TabIndex = 0;
            picImage.TabStop = false;
            // 
            // menuStrip1
            // 
            menuStrip1.ImageScalingSize = new Size(32, 32);
            menuStrip1.Items.AddRange(new ToolStripItem[] { fileToolStripMenuItem, versionToolStripMenuItem, sortToolStripMenuItem, dataDexToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(730, 24);
            menuStrip1.TabIndex = 1;
            menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { openToolStripMenuItem, saveToolStripMenuItem });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new Size(37, 20);
            fileToolStripMenuItem.Text = "&File";
            // 
            // openToolStripMenuItem
            // 
            openToolStripMenuItem.Name = "openToolStripMenuItem";
            openToolStripMenuItem.Size = new Size(103, 22);
            openToolStripMenuItem.Text = "&Open";
            openToolStripMenuItem.Click += openToolStripMenuItem_Click;
            // 
            // saveToolStripMenuItem
            // 
            saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            saveToolStripMenuItem.Size = new Size(103, 22);
            saveToolStripMenuItem.Text = "&Save";
            // 
            // versionToolStripMenuItem
            // 
            versionToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { trueRBYToolStripMenuItem, rBYWithTradesToolStripMenuItem, toolStripMenuItem1, trueGSCToolStripMenuItem, gSCWithTradesToolStripMenuItem, toolStripMenuItem2, ruSaOnlyToolStripMenuItem, fullAdvancedToolStripMenuItem, toolStripMenuItem3, noModLoadedToolStripMenuItem, viewCurrentModsToolStripMenuItem, loadNewModFileToolStripMenuItem });
            versionToolStripMenuItem.Name = "versionToolStripMenuItem";
            versionToolStripMenuItem.Size = new Size(57, 20);
            versionToolStripMenuItem.Text = "&Version";
            // 
            // trueRBYToolStripMenuItem
            // 
            trueRBYToolStripMenuItem.Name = "trueRBYToolStripMenuItem";
            trueRBYToolStripMenuItem.Size = new Size(185, 22);
            trueRBYToolStripMenuItem.Text = "True RBY";
            trueRBYToolStripMenuItem.Click += trueRBYToolStripMenuItem_Click;
            // 
            // rBYWithTradesToolStripMenuItem
            // 
            rBYWithTradesToolStripMenuItem.Name = "rBYWithTradesToolStripMenuItem";
            rBYWithTradesToolStripMenuItem.Size = new Size(185, 22);
            rBYWithTradesToolStripMenuItem.Text = "RBY with Trades";
            rBYWithTradesToolStripMenuItem.Click += rBYWithTradesToolStripMenuItem_Click;
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new Size(182, 6);
            // 
            // trueGSCToolStripMenuItem
            // 
            trueGSCToolStripMenuItem.Name = "trueGSCToolStripMenuItem";
            trueGSCToolStripMenuItem.Size = new Size(185, 22);
            trueGSCToolStripMenuItem.Text = "True GSC";
            trueGSCToolStripMenuItem.Click += trueGSCToolStripMenuItem_Click;
            // 
            // gSCWithTradesToolStripMenuItem
            // 
            gSCWithTradesToolStripMenuItem.Name = "gSCWithTradesToolStripMenuItem";
            gSCWithTradesToolStripMenuItem.Size = new Size(185, 22);
            gSCWithTradesToolStripMenuItem.Text = "GSC with Trades";
            gSCWithTradesToolStripMenuItem.Click += gSCWithTradesToolStripMenuItem_Click;
            // 
            // toolStripMenuItem2
            // 
            toolStripMenuItem2.Name = "toolStripMenuItem2";
            toolStripMenuItem2.Size = new Size(182, 6);
            // 
            // ruSaOnlyToolStripMenuItem
            // 
            ruSaOnlyToolStripMenuItem.Name = "ruSaOnlyToolStripMenuItem";
            ruSaOnlyToolStripMenuItem.Size = new Size(185, 22);
            ruSaOnlyToolStripMenuItem.Text = "Ru/Sa Only";
            ruSaOnlyToolStripMenuItem.Click += ruSaOnlyToolStripMenuItem_Click;
            // 
            // fullAdvancedToolStripMenuItem
            // 
            fullAdvancedToolStripMenuItem.Checked = true;
            fullAdvancedToolStripMenuItem.CheckState = CheckState.Checked;
            fullAdvancedToolStripMenuItem.Name = "fullAdvancedToolStripMenuItem";
            fullAdvancedToolStripMenuItem.Size = new Size(185, 22);
            fullAdvancedToolStripMenuItem.Text = "Full Advanced";
            fullAdvancedToolStripMenuItem.Click += fullAdvancedToolStripMenuItem_Click;
            // 
            // toolStripMenuItem3
            // 
            toolStripMenuItem3.Name = "toolStripMenuItem3";
            toolStripMenuItem3.Size = new Size(182, 6);
            // 
            // noModLoadedToolStripMenuItem
            // 
            noModLoadedToolStripMenuItem.Enabled = false;
            noModLoadedToolStripMenuItem.Name = "noModLoadedToolStripMenuItem";
            noModLoadedToolStripMenuItem.Size = new Size(185, 22);
            noModLoadedToolStripMenuItem.Text = "No Mod Loaded";
            // 
            // viewCurrentModsToolStripMenuItem
            // 
            viewCurrentModsToolStripMenuItem.Name = "viewCurrentModsToolStripMenuItem";
            viewCurrentModsToolStripMenuItem.Size = new Size(185, 22);
            viewCurrentModsToolStripMenuItem.Text = "View Current Mods...";
            // 
            // loadNewModFileToolStripMenuItem
            // 
            loadNewModFileToolStripMenuItem.Name = "loadNewModFileToolStripMenuItem";
            loadNewModFileToolStripMenuItem.Size = new Size(185, 22);
            loadNewModFileToolStripMenuItem.Text = "Load New Mod File...";
            // 
            // sortToolStripMenuItem
            // 
            sortToolStripMenuItem.Name = "sortToolStripMenuItem";
            sortToolStripMenuItem.Size = new Size(40, 20);
            sortToolStripMenuItem.Text = "&Sort";
            // 
            // dataDexToolStripMenuItem
            // 
            dataDexToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { pokedexToolStripMenuItem, moveDexToolStripMenuItem, battleDexToolStripMenuItem, damageCalcToolStripMenuItem });
            dataDexToolStripMenuItem.Name = "dataDexToolStripMenuItem";
            dataDexToolStripMenuItem.Size = new Size(63, 20);
            dataDexToolStripMenuItem.Text = "&DataDex";
            // 
            // pokedexToolStripMenuItem
            // 
            pokedexToolStripMenuItem.Name = "pokedexToolStripMenuItem";
            pokedexToolStripMenuItem.Size = new Size(141, 22);
            pokedexToolStripMenuItem.Text = "&Pokedex";
            // 
            // moveDexToolStripMenuItem
            // 
            moveDexToolStripMenuItem.Name = "moveDexToolStripMenuItem";
            moveDexToolStripMenuItem.Size = new Size(141, 22);
            moveDexToolStripMenuItem.Text = "&MoveDex";
            // 
            // battleDexToolStripMenuItem
            // 
            battleDexToolStripMenuItem.Name = "battleDexToolStripMenuItem";
            battleDexToolStripMenuItem.Size = new Size(141, 22);
            battleDexToolStripMenuItem.Text = "&BattleDex";
            // 
            // damageCalcToolStripMenuItem
            // 
            damageCalcToolStripMenuItem.Name = "damageCalcToolStripMenuItem";
            damageCalcToolStripMenuItem.Size = new Size(141, 22);
            damageCalcToolStripMenuItem.Text = "&DamageCalc";
            // 
            // ofdPnb
            // 
            ofdPnb.DefaultExt = "pnb";
            ofdPnb.FileName = "openFileDialog1";
            ofdPnb.Filter = "Netbattle Saves|*.pnb";
            ofdPnb.FileOk += ofdPnb_FileOk;
            // 
            // TeamBuilder
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(730, 645);
            Controls.Add(tabControl1);
            Controls.Add(menuStrip1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = menuStrip1;
            Margin = new Padding(6, 7, 6, 7);
            Name = "TeamBuilder";
            Text = "Team Builder";
            FormClosing += TeamBuilder_FormClosing;
            Load += TeamBuilder_Load;
            tabControl1.ResumeLayout(false);
            tbTrainer.ResumeLayout(false);
            grpAutoMessages.ResumeLayout(false);
            grpAutoMessages.PerformLayout();
            grpInfo.ResumeLayout(false);
            grpInfo.PerformLayout();
            grpImage.ResumeLayout(false);
            tbSlot1.ResumeLayout(false);
            tbSlot1.PerformLayout();
            grpStats.ResumeLayout(false);
            grpStats.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)picImage).EndInit();
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        private System.Windows.Forms.ToolStripMenuItem damageCalcToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem battleDexToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem pokedexToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem moveDexToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem sortToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem dataDexToolStripMenuItem;

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tbTrainer;
        private System.Windows.Forms.TabPage tbSlot1;
        private System.Windows.Forms.PictureBox picImage;
        private System.Windows.Forms.GroupBox grpAutoMessages;
        private System.Windows.Forms.TextBox txtLoseMessage;
        private System.Windows.Forms.TextBox txtWinMessage;
        private System.Windows.Forms.Label lblAutoLose;
        private System.Windows.Forms.Label lblAutoWin;
        private System.Windows.Forms.GroupBox grpInfo;
        private System.Windows.Forms.ComboBox dropGraphics;
        private System.Windows.Forms.TextBox txtExtraInfo;
        private System.Windows.Forms.TextBox txtUsername;
        private System.Windows.Forms.Label lblExtraInfo;
        private System.Windows.Forms.Label lblGraphics;
        private System.Windows.Forms.Label lblUsername;
        private System.Windows.Forms.GroupBox grpImage;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ComboBox dropPokemon;
        private System.Windows.Forms.Button btnSwitch;
        private System.Windows.Forms.Label lblPokemon;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ComboBox dropItem;
        private System.Windows.Forms.TextBox txtNickname;
        private System.Windows.Forms.TextBox txtMove4;
        private System.Windows.Forms.TextBox txtMove3;
        private System.Windows.Forms.TextBox txtMove2;
        private System.Windows.Forms.TextBox txtMove1;
        private System.Windows.Forms.Label lblLevel;
        private System.Windows.Forms.GroupBox grpStats;
        private System.Windows.Forms.Label lblSpecialDefense;
        private System.Windows.Forms.Label lblSpecialAttack;
        private System.Windows.Forms.Label lblSpeed;
        private System.Windows.Forms.Label lblDefense;
        private System.Windows.Forms.Label lblAttack;
        private System.Windows.Forms.Label lblHp;
        private System.Windows.Forms.Label lblTypes;
        private System.Windows.Forms.ListView listMoves;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblNickname;
        private System.Windows.Forms.ColumnHeader colMove;
        private System.Windows.Forms.ColumnHeader colPower;
        private System.Windows.Forms.ColumnHeader colAcc;
        private System.Windows.Forms.ColumnHeader colPP;
        private System.Windows.Forms.ColumnHeader colLearned;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.Button btnPoke6;
        private System.Windows.Forms.ImageList imageList2;
        private System.Windows.Forms.Button btnPoke5;
        private System.Windows.Forms.Button btnPoke4;
        private System.Windows.Forms.Button btnPoke3;
        private System.Windows.Forms.Button btnPoke2;
        private System.Windows.Forms.Button btnPoke1;
        private System.Windows.Forms.ToolStripMenuItem versionToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem trueRBYToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem rBYWithTradesToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem trueGSCToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem gSCWithTradesToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem ruSaOnlyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fullAdvancedToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem noModLoadedToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewCurrentModsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem loadNewModFileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.OpenFileDialog ofdPnb;
    }
}