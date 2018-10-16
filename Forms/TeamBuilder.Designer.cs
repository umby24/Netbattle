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
        private void InitializeComponent() {
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tbTrainer = new System.Windows.Forms.TabPage();
            this.grpAutoMessages = new System.Windows.Forms.GroupBox();
            this.txtLoseMessage = new System.Windows.Forms.TextBox();
            this.txtWinMessage = new System.Windows.Forms.TextBox();
            this.lblAutoLose = new System.Windows.Forms.Label();
            this.lblAutoWin = new System.Windows.Forms.Label();
            this.grpInfo = new System.Windows.Forms.GroupBox();
            this.dropGraphics = new System.Windows.Forms.ComboBox();
            this.txtExtraInfo = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.lblExtraInfo = new System.Windows.Forms.Label();
            this.lblGraphics = new System.Windows.Forms.Label();
            this.lblUsername = new System.Windows.Forms.Label();
            this.grpImage = new System.Windows.Forms.GroupBox();
            this.listView1 = new System.Windows.Forms.ListView();
            this.tbSlot1 = new System.Windows.Forms.TabPage();
            this.dropItem = new System.Windows.Forms.ComboBox();
            this.txtNickname = new System.Windows.Forms.TextBox();
            this.txtMove4 = new System.Windows.Forms.TextBox();
            this.txtMove3 = new System.Windows.Forms.TextBox();
            this.txtMove2 = new System.Windows.Forms.TextBox();
            this.txtMove1 = new System.Windows.Forms.TextBox();
            this.lblLevel = new System.Windows.Forms.Label();
            this.grpStats = new System.Windows.Forms.GroupBox();
            this.lblSpecialDefense = new System.Windows.Forms.Label();
            this.lblSpecialAttack = new System.Windows.Forms.Label();
            this.lblSpeed = new System.Windows.Forms.Label();
            this.lblDefense = new System.Windows.Forms.Label();
            this.lblAttack = new System.Windows.Forms.Label();
            this.lblHp = new System.Windows.Forms.Label();
            this.lblTypes = new System.Windows.Forms.Label();
            this.listMoves = new System.Windows.Forms.ListView();
            this.colMove = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colPower = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colAcc = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colPP = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colLearned = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.label2 = new System.Windows.Forms.Label();
            this.lblNickname = new System.Windows.Forms.Label();
            this.dropPokemon = new System.Windows.Forms.ComboBox();
            this.btnSwitch = new System.Windows.Forms.Button();
            this.lblPokemon = new System.Windows.Forms.Label();
            this.picImage = new System.Windows.Forms.PictureBox();
            this.tbSlot2 = new System.Windows.Forms.TabPage();
            this.tbSlot3 = new System.Windows.Forms.TabPage();
            this.tbSlot4 = new System.Windows.Forms.TabPage();
            this.tbSlot5 = new System.Windows.Forms.TabPage();
            this.tbSlot6 = new System.Windows.Forms.TabPage();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.tabControl1.SuspendLayout();
            this.tbTrainer.SuspendLayout();
            this.grpAutoMessages.SuspendLayout();
            this.grpInfo.SuspendLayout();
            this.grpImage.SuspendLayout();
            this.tbSlot1.SuspendLayout();
            this.grpStats.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picImage)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tbTrainer);
            this.tabControl1.Controls.Add(this.tbSlot1);
            this.tabControl1.Controls.Add(this.tbSlot2);
            this.tabControl1.Controls.Add(this.tbSlot3);
            this.tabControl1.Controls.Add(this.tbSlot4);
            this.tabControl1.Controls.Add(this.tbSlot5);
            this.tabControl1.Controls.Add(this.tbSlot6);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 24);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(800, 426);
            this.tabControl1.TabIndex = 0;
            // 
            // tbTrainer
            // 
            this.tbTrainer.Controls.Add(this.grpAutoMessages);
            this.tbTrainer.Controls.Add(this.grpInfo);
            this.tbTrainer.Controls.Add(this.grpImage);
            this.tbTrainer.Location = new System.Drawing.Point(4, 22);
            this.tbTrainer.Name = "tbTrainer";
            this.tbTrainer.Padding = new System.Windows.Forms.Padding(3);
            this.tbTrainer.Size = new System.Drawing.Size(792, 400);
            this.tbTrainer.TabIndex = 0;
            this.tbTrainer.Text = "Trainer Info";
            this.tbTrainer.UseVisualStyleBackColor = true;
            // 
            // grpAutoMessages
            // 
            this.grpAutoMessages.Controls.Add(this.txtLoseMessage);
            this.grpAutoMessages.Controls.Add(this.txtWinMessage);
            this.grpAutoMessages.Controls.Add(this.lblAutoLose);
            this.grpAutoMessages.Controls.Add(this.lblAutoWin);
            this.grpAutoMessages.Location = new System.Drawing.Point(164, 262);
            this.grpAutoMessages.Name = "grpAutoMessages";
            this.grpAutoMessages.Size = new System.Drawing.Size(620, 154);
            this.grpAutoMessages.TabIndex = 3;
            this.grpAutoMessages.TabStop = false;
            this.grpAutoMessages.Text = "Auto Messages";
            // 
            // txtLoseMessage
            // 
            this.txtLoseMessage.Location = new System.Drawing.Point(19, 97);
            this.txtLoseMessage.Multiline = true;
            this.txtLoseMessage.Name = "txtLoseMessage";
            this.txtLoseMessage.Size = new System.Drawing.Size(581, 46);
            this.txtLoseMessage.TabIndex = 6;
            // 
            // txtWinMessage
            // 
            this.txtWinMessage.Location = new System.Drawing.Point(19, 32);
            this.txtWinMessage.Multiline = true;
            this.txtWinMessage.Name = "txtWinMessage";
            this.txtWinMessage.Size = new System.Drawing.Size(581, 46);
            this.txtWinMessage.TabIndex = 5;
            // 
            // lblAutoLose
            // 
            this.lblAutoLose.AutoSize = true;
            this.lblAutoLose.Location = new System.Drawing.Point(18, 81);
            this.lblAutoLose.Name = "lblAutoLose";
            this.lblAutoLose.Size = new System.Drawing.Size(30, 13);
            this.lblAutoLose.TabIndex = 1;
            this.lblAutoLose.Text = "Lose";
            // 
            // lblAutoWin
            // 
            this.lblAutoWin.AutoSize = true;
            this.lblAutoWin.Location = new System.Drawing.Point(22, 16);
            this.lblAutoWin.Name = "lblAutoWin";
            this.lblAutoWin.Size = new System.Drawing.Size(26, 13);
            this.lblAutoWin.TabIndex = 0;
            this.lblAutoWin.Text = "Win";
            // 
            // grpInfo
            // 
            this.grpInfo.Controls.Add(this.dropGraphics);
            this.grpInfo.Controls.Add(this.txtExtraInfo);
            this.grpInfo.Controls.Add(this.textBox1);
            this.grpInfo.Controls.Add(this.lblExtraInfo);
            this.grpInfo.Controls.Add(this.lblGraphics);
            this.grpInfo.Controls.Add(this.lblUsername);
            this.grpInfo.Location = new System.Drawing.Point(164, 6);
            this.grpInfo.Name = "grpInfo";
            this.grpInfo.Size = new System.Drawing.Size(620, 190);
            this.grpInfo.TabIndex = 2;
            this.grpInfo.TabStop = false;
            this.grpInfo.Text = "Info";
            // 
            // dropGraphics
            // 
            this.dropGraphics.FormattingEnabled = true;
            this.dropGraphics.Items.AddRange(new object[] {
            "Green",
            "Red/Blue",
            "Yellow",
            "Gold",
            "Silver",
            "Ruby/Sapphire",
            "Leaf/Fire",
            "Emerald"});
            this.dropGraphics.Location = new System.Drawing.Point(77, 46);
            this.dropGraphics.Name = "dropGraphics";
            this.dropGraphics.Size = new System.Drawing.Size(134, 21);
            this.dropGraphics.TabIndex = 5;
            this.dropGraphics.SelectedIndexChanged += new System.EventHandler(this.dropGraphics_SelectedIndexChanged);
            // 
            // txtExtraInfo
            // 
            this.txtExtraInfo.Location = new System.Drawing.Point(19, 97);
            this.txtExtraInfo.MaxLength = 200;
            this.txtExtraInfo.Multiline = true;
            this.txtExtraInfo.Name = "txtExtraInfo";
            this.txtExtraInfo.Size = new System.Drawing.Size(595, 73);
            this.txtExtraInfo.TabIndex = 4;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(77, 13);
            this.textBox1.MaxLength = 20;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(134, 20);
            this.textBox1.TabIndex = 3;
            // 
            // lblExtraInfo
            // 
            this.lblExtraInfo.AutoSize = true;
            this.lblExtraInfo.Location = new System.Drawing.Point(19, 81);
            this.lblExtraInfo.Name = "lblExtraInfo";
            this.lblExtraInfo.Size = new System.Drawing.Size(52, 13);
            this.lblExtraInfo.TabIndex = 2;
            this.lblExtraInfo.Text = "Extra Info";
            // 
            // lblGraphics
            // 
            this.lblGraphics.AutoSize = true;
            this.lblGraphics.Location = new System.Drawing.Point(22, 49);
            this.lblGraphics.Name = "lblGraphics";
            this.lblGraphics.Size = new System.Drawing.Size(49, 13);
            this.lblGraphics.TabIndex = 1;
            this.lblGraphics.Text = "Graphics";
            // 
            // lblUsername
            // 
            this.lblUsername.AutoSize = true;
            this.lblUsername.Location = new System.Drawing.Point(16, 16);
            this.lblUsername.Name = "lblUsername";
            this.lblUsername.Size = new System.Drawing.Size(55, 13);
            this.lblUsername.TabIndex = 0;
            this.lblUsername.Text = "Username";
            // 
            // grpImage
            // 
            this.grpImage.Controls.Add(this.listView1);
            this.grpImage.Location = new System.Drawing.Point(8, 6);
            this.grpImage.Name = "grpImage";
            this.grpImage.Size = new System.Drawing.Size(150, 410);
            this.grpImage.TabIndex = 1;
            this.grpImage.TabStop = false;
            this.grpImage.Text = "Image";
            // 
            // listView1
            // 
            this.listView1.Location = new System.Drawing.Point(6, 19);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(138, 380);
            this.listView1.TabIndex = 7;
            this.listView1.UseCompatibleStateImageBehavior = false;
            // 
            // tbSlot1
            // 
            this.tbSlot1.Controls.Add(this.dropItem);
            this.tbSlot1.Controls.Add(this.txtNickname);
            this.tbSlot1.Controls.Add(this.txtMove4);
            this.tbSlot1.Controls.Add(this.txtMove3);
            this.tbSlot1.Controls.Add(this.txtMove2);
            this.tbSlot1.Controls.Add(this.txtMove1);
            this.tbSlot1.Controls.Add(this.lblLevel);
            this.tbSlot1.Controls.Add(this.grpStats);
            this.tbSlot1.Controls.Add(this.listMoves);
            this.tbSlot1.Controls.Add(this.label2);
            this.tbSlot1.Controls.Add(this.lblNickname);
            this.tbSlot1.Controls.Add(this.dropPokemon);
            this.tbSlot1.Controls.Add(this.btnSwitch);
            this.tbSlot1.Controls.Add(this.lblPokemon);
            this.tbSlot1.Controls.Add(this.picImage);
            this.tbSlot1.Location = new System.Drawing.Point(4, 22);
            this.tbSlot1.Name = "tbSlot1";
            this.tbSlot1.Padding = new System.Windows.Forms.Padding(3);
            this.tbSlot1.Size = new System.Drawing.Size(792, 400);
            this.tbSlot1.TabIndex = 1;
            this.tbSlot1.Text = "Slot 1";
            this.tbSlot1.UseVisualStyleBackColor = true;
            // 
            // dropItem
            // 
            this.dropItem.FormattingEnabled = true;
            this.dropItem.Location = new System.Drawing.Point(221, 56);
            this.dropItem.Name = "dropItem";
            this.dropItem.Size = new System.Drawing.Size(121, 21);
            this.dropItem.TabIndex = 19;
            // 
            // txtNickname
            // 
            this.txtNickname.Location = new System.Drawing.Point(221, 30);
            this.txtNickname.Name = "txtNickname";
            this.txtNickname.Size = new System.Drawing.Size(121, 20);
            this.txtNickname.TabIndex = 18;
            // 
            // txtMove4
            // 
            this.txtMove4.Location = new System.Drawing.Point(373, 361);
            this.txtMove4.Name = "txtMove4";
            this.txtMove4.Size = new System.Drawing.Size(100, 20);
            this.txtMove4.TabIndex = 17;
            // 
            // txtMove3
            // 
            this.txtMove3.Location = new System.Drawing.Point(253, 361);
            this.txtMove3.Name = "txtMove3";
            this.txtMove3.Size = new System.Drawing.Size(100, 20);
            this.txtMove3.TabIndex = 16;
            // 
            // txtMove2
            // 
            this.txtMove2.Location = new System.Drawing.Point(132, 361);
            this.txtMove2.Name = "txtMove2";
            this.txtMove2.Size = new System.Drawing.Size(100, 20);
            this.txtMove2.TabIndex = 15;
            // 
            // txtMove1
            // 
            this.txtMove1.Location = new System.Drawing.Point(9, 361);
            this.txtMove1.Name = "txtMove1";
            this.txtMove1.Size = new System.Drawing.Size(100, 20);
            this.txtMove1.TabIndex = 14;
            // 
            // lblLevel
            // 
            this.lblLevel.AutoSize = true;
            this.lblLevel.Location = new System.Drawing.Point(33, 111);
            this.lblLevel.Name = "lblLevel";
            this.lblLevel.Size = new System.Drawing.Size(72, 13);
            this.lblLevel.TabIndex = 13;
            this.lblLevel.Text = "Lv. 100, Male";
            // 
            // grpStats
            // 
            this.grpStats.Controls.Add(this.lblSpecialDefense);
            this.grpStats.Controls.Add(this.lblSpecialAttack);
            this.grpStats.Controls.Add(this.lblSpeed);
            this.grpStats.Controls.Add(this.lblDefense);
            this.grpStats.Controls.Add(this.lblAttack);
            this.grpStats.Controls.Add(this.lblHp);
            this.grpStats.Controls.Add(this.lblTypes);
            this.grpStats.Location = new System.Drawing.Point(429, 6);
            this.grpStats.Name = "grpStats";
            this.grpStats.Size = new System.Drawing.Size(355, 129);
            this.grpStats.TabIndex = 7;
            this.grpStats.TabStop = false;
            this.grpStats.Text = "Stats";
            // 
            // lblSpecialDefense
            // 
            this.lblSpecialDefense.AutoSize = true;
            this.lblSpecialDefense.Location = new System.Drawing.Point(6, 105);
            this.lblSpecialDefense.Name = "lblSpecialDefense";
            this.lblSpecialDefense.Size = new System.Drawing.Size(66, 13);
            this.lblSpecialDefense.TabIndex = 11;
            this.lblSpecialDefense.Text = "Sp. Defense";
            // 
            // lblSpecialAttack
            // 
            this.lblSpecialAttack.AutoSize = true;
            this.lblSpecialAttack.Location = new System.Drawing.Point(6, 88);
            this.lblSpecialAttack.Name = "lblSpecialAttack";
            this.lblSpecialAttack.Size = new System.Drawing.Size(57, 13);
            this.lblSpecialAttack.TabIndex = 10;
            this.lblSpecialAttack.Text = "Sp. Attack";
            // 
            // lblSpeed
            // 
            this.lblSpeed.AutoSize = true;
            this.lblSpeed.Location = new System.Drawing.Point(6, 72);
            this.lblSpeed.Name = "lblSpeed";
            this.lblSpeed.Size = new System.Drawing.Size(38, 13);
            this.lblSpeed.TabIndex = 9;
            this.lblSpeed.Text = "Speed";
            // 
            // lblDefense
            // 
            this.lblDefense.AutoSize = true;
            this.lblDefense.Location = new System.Drawing.Point(6, 52);
            this.lblDefense.Name = "lblDefense";
            this.lblDefense.Size = new System.Drawing.Size(47, 13);
            this.lblDefense.TabIndex = 8;
            this.lblDefense.Text = "Defense";
            // 
            // lblAttack
            // 
            this.lblAttack.AutoSize = true;
            this.lblAttack.Location = new System.Drawing.Point(6, 34);
            this.lblAttack.Name = "lblAttack";
            this.lblAttack.Size = new System.Drawing.Size(38, 13);
            this.lblAttack.TabIndex = 7;
            this.lblAttack.Text = "Attack";
            // 
            // lblHp
            // 
            this.lblHp.AutoSize = true;
            this.lblHp.Location = new System.Drawing.Point(6, 16);
            this.lblHp.Name = "lblHp";
            this.lblHp.Size = new System.Drawing.Size(22, 13);
            this.lblHp.TabIndex = 6;
            this.lblHp.Text = "HP";
            // 
            // lblTypes
            // 
            this.lblTypes.AutoSize = true;
            this.lblTypes.Location = new System.Drawing.Point(180, 16);
            this.lblTypes.Name = "lblTypes";
            this.lblTypes.Size = new System.Drawing.Size(34, 13);
            this.lblTypes.TabIndex = 12;
            this.lblTypes.Text = "Type:";
            // 
            // listMoves
            // 
            this.listMoves.CheckBoxes = true;
            this.listMoves.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colMove,
            this.colPower,
            this.colAcc,
            this.colPP,
            this.colLearned});
            this.listMoves.FullRowSelect = true;
            this.listMoves.Location = new System.Drawing.Point(9, 141);
            this.listMoves.Name = "listMoves";
            this.listMoves.Size = new System.Drawing.Size(775, 214);
            this.listMoves.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.listMoves.TabIndex = 6;
            this.listMoves.UseCompatibleStateImageBehavior = false;
            this.listMoves.View = System.Windows.Forms.View.Details;
            // 
            // colMove
            // 
            this.colMove.Text = "Move";
            this.colMove.Width = 237;
            // 
            // colPower
            // 
            this.colPower.Text = "Power";
            this.colPower.Width = 93;
            // 
            // colAcc
            // 
            this.colAcc.Text = "Acc.";
            this.colAcc.Width = 112;
            // 
            // colPP
            // 
            this.colPP.Text = "PP";
            // 
            // colLearned
            // 
            this.colLearned.Text = "Learned By";
            this.colLearned.Width = 157;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(188, 59);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(27, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Item";
            // 
            // lblNickname
            // 
            this.lblNickname.AutoSize = true;
            this.lblNickname.Location = new System.Drawing.Point(160, 35);
            this.lblNickname.Name = "lblNickname";
            this.lblNickname.Size = new System.Drawing.Size(55, 13);
            this.lblNickname.TabIndex = 4;
            this.lblNickname.Text = "Nickname";
            // 
            // dropPokemon
            // 
            this.dropPokemon.FormattingEnabled = true;
            this.dropPokemon.Location = new System.Drawing.Point(221, 3);
            this.dropPokemon.Name = "dropPokemon";
            this.dropPokemon.Size = new System.Drawing.Size(121, 21);
            this.dropPokemon.TabIndex = 3;
            this.dropPokemon.SelectedIndexChanged += new System.EventHandler(this.dropPokemon_SelectedIndexChanged);
            // 
            // btnSwitch
            // 
            this.btnSwitch.Location = new System.Drawing.Point(348, 1);
            this.btnSwitch.Name = "btnSwitch";
            this.btnSwitch.Size = new System.Drawing.Size(75, 23);
            this.btnSwitch.TabIndex = 2;
            this.btnSwitch.Text = "Expert";
            this.btnSwitch.UseVisualStyleBackColor = true;
            // 
            // lblPokemon
            // 
            this.lblPokemon.AutoSize = true;
            this.lblPokemon.Location = new System.Drawing.Point(163, 8);
            this.lblPokemon.Name = "lblPokemon";
            this.lblPokemon.Size = new System.Drawing.Size(52, 13);
            this.lblPokemon.TabIndex = 1;
            this.lblPokemon.Text = "Pokemon";
            // 
            // picImage
            // 
            this.picImage.Location = new System.Drawing.Point(8, 8);
            this.picImage.Name = "picImage";
            this.picImage.Size = new System.Drawing.Size(122, 100);
            this.picImage.TabIndex = 0;
            this.picImage.TabStop = false;
            // 
            // tbSlot2
            // 
            this.tbSlot2.Location = new System.Drawing.Point(4, 22);
            this.tbSlot2.Name = "tbSlot2";
            this.tbSlot2.Size = new System.Drawing.Size(792, 400);
            this.tbSlot2.TabIndex = 2;
            this.tbSlot2.Text = "Slot 2";
            this.tbSlot2.UseVisualStyleBackColor = true;
            // 
            // tbSlot3
            // 
            this.tbSlot3.Location = new System.Drawing.Point(4, 22);
            this.tbSlot3.Name = "tbSlot3";
            this.tbSlot3.Size = new System.Drawing.Size(792, 400);
            this.tbSlot3.TabIndex = 3;
            this.tbSlot3.Text = "Slot 3";
            this.tbSlot3.UseVisualStyleBackColor = true;
            // 
            // tbSlot4
            // 
            this.tbSlot4.Location = new System.Drawing.Point(4, 22);
            this.tbSlot4.Name = "tbSlot4";
            this.tbSlot4.Size = new System.Drawing.Size(792, 400);
            this.tbSlot4.TabIndex = 4;
            this.tbSlot4.Text = "Slot 4";
            this.tbSlot4.UseVisualStyleBackColor = true;
            // 
            // tbSlot5
            // 
            this.tbSlot5.Location = new System.Drawing.Point(4, 22);
            this.tbSlot5.Name = "tbSlot5";
            this.tbSlot5.Size = new System.Drawing.Size(792, 400);
            this.tbSlot5.TabIndex = 5;
            this.tbSlot5.Text = "Slot 5";
            this.tbSlot5.UseVisualStyleBackColor = true;
            // 
            // tbSlot6
            // 
            this.tbSlot6.Location = new System.Drawing.Point(4, 22);
            this.tbSlot6.Name = "tbSlot6";
            this.tbSlot6.Size = new System.Drawing.Size(792, 400);
            this.tbSlot6.TabIndex = 6;
            this.tbSlot6.Text = "Slot 6";
            this.tbSlot6.UseVisualStyleBackColor = true;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(800, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // TeamBuilder
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "TeamBuilder";
            this.Text = "Team Builder";
            this.Load += new System.EventHandler(this.TeamBuilder_Load);
            this.tabControl1.ResumeLayout(false);
            this.tbTrainer.ResumeLayout(false);
            this.grpAutoMessages.ResumeLayout(false);
            this.grpAutoMessages.PerformLayout();
            this.grpInfo.ResumeLayout(false);
            this.grpInfo.PerformLayout();
            this.grpImage.ResumeLayout(false);
            this.tbSlot1.ResumeLayout(false);
            this.tbSlot1.PerformLayout();
            this.grpStats.ResumeLayout(false);
            this.grpStats.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picImage)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tbTrainer;
        private System.Windows.Forms.TabPage tbSlot1;
        private System.Windows.Forms.TabPage tbSlot2;
        private System.Windows.Forms.TabPage tbSlot3;
        private System.Windows.Forms.TabPage tbSlot4;
        private System.Windows.Forms.TabPage tbSlot5;
        private System.Windows.Forms.TabPage tbSlot6;
        private System.Windows.Forms.PictureBox picImage;
        private System.Windows.Forms.GroupBox grpAutoMessages;
        private System.Windows.Forms.TextBox txtLoseMessage;
        private System.Windows.Forms.TextBox txtWinMessage;
        private System.Windows.Forms.Label lblAutoLose;
        private System.Windows.Forms.Label lblAutoWin;
        private System.Windows.Forms.GroupBox grpInfo;
        private System.Windows.Forms.ComboBox dropGraphics;
        private System.Windows.Forms.TextBox txtExtraInfo;
        private System.Windows.Forms.TextBox textBox1;
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
    }
}