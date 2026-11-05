using System.ComponentModel;

namespace Netbattle.Forms
{
    partial class Settings
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Settings));
            lblRegAddr = new System.Windows.Forms.Label();
            tabControl1 = new System.Windows.Forms.TabControl();
            tabPage1 = new System.Windows.Forms.TabPage();
            btnStoredPassword = new System.Windows.Forms.Button();
            groupBox3 = new System.Windows.Forms.GroupBox();
            chkReplayAutosave = new System.Windows.Forms.CheckBox();
            chkReplays = new System.Windows.Forms.CheckBox();
            groupBox4 = new System.Windows.Forms.GroupBox();
            txtRegAddr = new System.Windows.Forms.TextBox();
            groupBox2 = new System.Windows.Forms.GroupBox();
            chkLogsAutosave = new System.Windows.Forms.CheckBox();
            chkLogsSave = new System.Windows.Forms.CheckBox();
            groupBox1 = new System.Windows.Forms.GroupBox();
            rdoHideTeam = new System.Windows.Forms.RadioButton();
            rdoShowTeam = new System.Windows.Forms.RadioButton();
            chkAutoLoad = new System.Windows.Forms.CheckBox();
            tabPage2 = new System.Windows.Forms.TabPage();
            groupBox8 = new System.Windows.Forms.GroupBox();
            chkMusic = new System.Windows.Forms.CheckBox();
            chkSounds = new System.Windows.Forms.CheckBox();
            groupBox7 = new System.Windows.Forms.GroupBox();
            chkAnimateHP = new System.Windows.Forms.CheckBox();
            chkEnableBackground = new System.Windows.Forms.CheckBox();
            groupBox6 = new System.Windows.Forms.GroupBox();
            rdoExtended = new System.Windows.Forms.RadioButton();
            rdoNormal = new System.Windows.Forms.RadioButton();
            chkTrainerPrefix = new System.Windows.Forms.CheckBox();
            chkUseNicks = new System.Windows.Forms.CheckBox();
            groupBox5 = new System.Windows.Forms.GroupBox();
            tabPage3 = new System.Windows.Forms.TabPage();
            label1 = new System.Windows.Forms.Label();
            txtTSFormat = new System.Windows.Forms.TextBox();
            chkTimestamp = new System.Windows.Forms.CheckBox();
            chkParseUrls = new System.Windows.Forms.CheckBox();
            groupBox9 = new System.Windows.Forms.GroupBox();
            chkShowNickChange = new System.Windows.Forms.CheckBox();
            chkReturnFromBattle = new System.Windows.Forms.CheckBox();
            chkShowTeamChange = new System.Windows.Forms.CheckBox();
            chkBattleResults = new System.Windows.Forms.CheckBox();
            chkBattleStart = new System.Windows.Forms.CheckBox();
            chkShowIgnore = new System.Windows.Forms.CheckBox();
            chkPlayerSignOff = new System.Windows.Forms.CheckBox();
            chkPlayerSignOn = new System.Windows.Forms.CheckBox();
            chkShowBack = new System.Windows.Forms.CheckBox();
            chkShowAway = new System.Windows.Forms.CheckBox();
            btnOk = new System.Windows.Forms.Button();
            btnCancel = new System.Windows.Forms.Button();
            tabControl1.SuspendLayout();
            tabPage1.SuspendLayout();
            groupBox3.SuspendLayout();
            groupBox4.SuspendLayout();
            groupBox2.SuspendLayout();
            groupBox1.SuspendLayout();
            tabPage2.SuspendLayout();
            groupBox8.SuspendLayout();
            groupBox7.SuspendLayout();
            groupBox6.SuspendLayout();
            tabPage3.SuspendLayout();
            groupBox9.SuspendLayout();
            SuspendLayout();
            // 
            // lblRegAddr
            // 
            lblRegAddr.Location = new System.Drawing.Point(7, 19);
            lblRegAddr.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            lblRegAddr.Name = "lblRegAddr";
            lblRegAddr.Size = new System.Drawing.Size(110, 17);
            lblRegAddr.TabIndex = 0;
            lblRegAddr.Text = "Registry Address:";
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(tabPage1);
            tabControl1.Controls.Add(tabPage2);
            tabControl1.Controls.Add(tabPage3);
            tabControl1.Dock = System.Windows.Forms.DockStyle.Top;
            tabControl1.Location = new System.Drawing.Point(0, 0);
            tabControl1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new System.Drawing.Size(340, 343);
            tabControl1.TabIndex = 1;
            tabControl1.SelectedIndexChanged += tabControl1_SelectedIndexChanged;
            // 
            // tabPage1
            // 
            tabPage1.Controls.Add(btnStoredPassword);
            tabPage1.Controls.Add(groupBox3);
            tabPage1.Controls.Add(groupBox4);
            tabPage1.Controls.Add(groupBox2);
            tabPage1.Controls.Add(groupBox1);
            tabPage1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)0));
            tabPage1.Location = new System.Drawing.Point(4, 24);
            tabPage1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            tabPage1.Size = new System.Drawing.Size(332, 315);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "System";
            tabPage1.UseVisualStyleBackColor = true;
            // 
            // btnStoredPassword
            // 
            btnStoredPassword.Location = new System.Drawing.Point(170, 220);
            btnStoredPassword.Name = "btnStoredPassword";
            btnStoredPassword.Size = new System.Drawing.Size(154, 36);
            btnStoredPassword.TabIndex = 6;
            btnStoredPassword.Text = "Change stored password";
            btnStoredPassword.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(chkReplayAutosave);
            groupBox3.Controls.Add(chkReplays);
            groupBox3.Location = new System.Drawing.Point(8, 179);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new System.Drawing.Size(135, 77);
            groupBox3.TabIndex = 4;
            groupBox3.TabStop = false;
            groupBox3.Text = "Replays";
            // 
            // chkReplayAutosave
            // 
            chkReplayAutosave.Location = new System.Drawing.Point(6, 45);
            chkReplayAutosave.Name = "chkReplayAutosave";
            chkReplayAutosave.Size = new System.Drawing.Size(122, 19);
            chkReplayAutosave.TabIndex = 3;
            chkReplayAutosave.Text = "Autosave";
            chkReplayAutosave.UseVisualStyleBackColor = true;
            // 
            // chkReplays
            // 
            chkReplays.Location = new System.Drawing.Point(6, 20);
            chkReplays.Name = "chkReplays";
            chkReplays.Size = new System.Drawing.Size(122, 19);
            chkReplays.TabIndex = 2;
            chkReplays.Text = "Prompt to save";
            chkReplays.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            groupBox4.Controls.Add(lblRegAddr);
            groupBox4.Controls.Add(txtRegAddr);
            groupBox4.Location = new System.Drawing.Point(154, 6);
            groupBox4.Name = "groupBox4";
            groupBox4.Size = new System.Drawing.Size(172, 82);
            groupBox4.TabIndex = 5;
            groupBox4.TabStop = false;
            groupBox4.Text = "Advanced";
            // 
            // txtRegAddr
            // 
            txtRegAddr.Location = new System.Drawing.Point(7, 39);
            txtRegAddr.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            txtRegAddr.Name = "txtRegAddr";
            txtRegAddr.Size = new System.Drawing.Size(148, 20);
            txtRegAddr.TabIndex = 1;
            txtRegAddr.TextChanged += txtRegAddr_TextChanged;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(chkLogsAutosave);
            groupBox2.Controls.Add(chkLogsSave);
            groupBox2.Location = new System.Drawing.Point(8, 94);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new System.Drawing.Size(135, 79);
            groupBox2.TabIndex = 3;
            groupBox2.TabStop = false;
            groupBox2.Text = "Logs";
            // 
            // chkLogsAutosave
            // 
            chkLogsAutosave.Location = new System.Drawing.Point(6, 44);
            chkLogsAutosave.Name = "chkLogsAutosave";
            chkLogsAutosave.Size = new System.Drawing.Size(122, 19);
            chkLogsAutosave.TabIndex = 1;
            chkLogsAutosave.Text = "Autosave";
            chkLogsAutosave.UseVisualStyleBackColor = true;
            // 
            // chkLogsSave
            // 
            chkLogsSave.Location = new System.Drawing.Point(6, 19);
            chkLogsSave.Name = "chkLogsSave";
            chkLogsSave.Size = new System.Drawing.Size(122, 19);
            chkLogsSave.TabIndex = 0;
            chkLogsSave.Text = "Prompt to save";
            chkLogsSave.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(rdoHideTeam);
            groupBox1.Controls.Add(rdoShowTeam);
            groupBox1.Controls.Add(chkAutoLoad);
            groupBox1.Location = new System.Drawing.Point(8, 6);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new System.Drawing.Size(135, 82);
            groupBox1.TabIndex = 2;
            groupBox1.TabStop = false;
            groupBox1.Text = "Team";
            // 
            // rdoHideTeam
            // 
            rdoHideTeam.Location = new System.Drawing.Point(6, 34);
            rdoHideTeam.Name = "rdoHideTeam";
            rdoHideTeam.Size = new System.Drawing.Size(123, 21);
            rdoHideTeam.TabIndex = 3;
            rdoHideTeam.TabStop = true;
            rdoHideTeam.Text = "Hide team from others";
            rdoHideTeam.UseVisualStyleBackColor = true;
            // 
            // rdoShowTeam
            // 
            rdoShowTeam.Location = new System.Drawing.Point(6, 15);
            rdoShowTeam.Name = "rdoShowTeam";
            rdoShowTeam.Size = new System.Drawing.Size(123, 21);
            rdoShowTeam.TabIndex = 2;
            rdoShowTeam.TabStop = true;
            rdoShowTeam.Text = "Show team to others";
            rdoShowTeam.UseVisualStyleBackColor = true;
            // 
            // chkAutoLoad
            // 
            chkAutoLoad.Location = new System.Drawing.Point(6, 57);
            chkAutoLoad.Name = "chkAutoLoad";
            chkAutoLoad.Size = new System.Drawing.Size(122, 19);
            chkAutoLoad.TabIndex = 1;
            chkAutoLoad.Text = "Auto-load on startup";
            chkAutoLoad.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            tabPage2.Controls.Add(groupBox8);
            tabPage2.Controls.Add(groupBox7);
            tabPage2.Controls.Add(groupBox6);
            tabPage2.Controls.Add(groupBox5);
            tabPage2.Location = new System.Drawing.Point(4, 24);
            tabPage2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            tabPage2.Size = new System.Drawing.Size(332, 315);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "Display & Sound";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox8
            // 
            groupBox8.Controls.Add(chkMusic);
            groupBox8.Controls.Add(chkSounds);
            groupBox8.Location = new System.Drawing.Point(169, 132);
            groupBox8.Name = "groupBox8";
            groupBox8.Size = new System.Drawing.Size(140, 100);
            groupBox8.TabIndex = 3;
            groupBox8.TabStop = false;
            groupBox8.Text = "Audio";
            // 
            // chkMusic
            // 
            chkMusic.Location = new System.Drawing.Point(18, 56);
            chkMusic.Name = "chkMusic";
            chkMusic.Size = new System.Drawing.Size(73, 19);
            chkMusic.TabIndex = 1;
            chkMusic.Text = "Music";
            chkMusic.UseVisualStyleBackColor = true;
            // 
            // chkSounds
            // 
            chkSounds.Location = new System.Drawing.Point(18, 31);
            chkSounds.Name = "chkSounds";
            chkSounds.Size = new System.Drawing.Size(73, 19);
            chkSounds.TabIndex = 0;
            chkSounds.Text = "Sounds";
            chkSounds.UseVisualStyleBackColor = true;
            // 
            // groupBox7
            // 
            groupBox7.Controls.Add(chkAnimateHP);
            groupBox7.Controls.Add(chkEnableBackground);
            groupBox7.Location = new System.Drawing.Point(169, 14);
            groupBox7.Name = "groupBox7";
            groupBox7.Size = new System.Drawing.Size(140, 100);
            groupBox7.TabIndex = 2;
            groupBox7.TabStop = false;
            groupBox7.Text = "Animation Display";
            // 
            // chkAnimateHP
            // 
            chkAnimateHP.Location = new System.Drawing.Point(6, 47);
            chkAnimateHP.Name = "chkAnimateHP";
            chkAnimateHP.Size = new System.Drawing.Size(128, 19);
            chkAnimateHP.TabIndex = 3;
            chkAnimateHP.Text = "Animate HP Bars";
            chkAnimateHP.UseVisualStyleBackColor = true;
            // 
            // chkEnableBackground
            // 
            chkEnableBackground.Location = new System.Drawing.Point(6, 22);
            chkEnableBackground.Name = "chkEnableBackground";
            chkEnableBackground.Size = new System.Drawing.Size(128, 19);
            chkEnableBackground.TabIndex = 2;
            chkEnableBackground.Text = "Enable Background";
            chkEnableBackground.UseVisualStyleBackColor = true;
            // 
            // groupBox6
            // 
            groupBox6.Controls.Add(rdoExtended);
            groupBox6.Controls.Add(rdoNormal);
            groupBox6.Controls.Add(chkTrainerPrefix);
            groupBox6.Controls.Add(chkUseNicks);
            groupBox6.Location = new System.Drawing.Point(11, 132);
            groupBox6.Name = "groupBox6";
            groupBox6.Size = new System.Drawing.Size(140, 123);
            groupBox6.TabIndex = 1;
            groupBox6.TabStop = false;
            groupBox6.Text = "Battle Messages";
            // 
            // rdoExtended
            // 
            rdoExtended.Location = new System.Drawing.Point(6, 40);
            rdoExtended.Name = "rdoExtended";
            rdoExtended.Size = new System.Drawing.Size(135, 27);
            rdoExtended.TabIndex = 4;
            rdoExtended.TabStop = true;
            rdoExtended.Text = "Extended (Log Style)";
            rdoExtended.UseVisualStyleBackColor = true;
            // 
            // rdoNormal
            // 
            rdoNormal.Location = new System.Drawing.Point(6, 22);
            rdoNormal.Name = "rdoNormal";
            rdoNormal.Size = new System.Drawing.Size(118, 17);
            rdoNormal.TabIndex = 3;
            rdoNormal.TabStop = true;
            rdoNormal.Text = "Normal";
            rdoNormal.UseVisualStyleBackColor = true;
            // 
            // chkTrainerPrefix
            // 
            chkTrainerPrefix.Location = new System.Drawing.Point(6, 98);
            chkTrainerPrefix.Name = "chkTrainerPrefix";
            chkTrainerPrefix.Size = new System.Drawing.Size(135, 19);
            chkTrainerPrefix.TabIndex = 2;
            chkTrainerPrefix.Text = "Trainer Name Prefix";
            chkTrainerPrefix.UseVisualStyleBackColor = true;
            // 
            // chkUseNicks
            // 
            chkUseNicks.Location = new System.Drawing.Point(6, 73);
            chkUseNicks.Name = "chkUseNicks";
            chkUseNicks.Size = new System.Drawing.Size(111, 19);
            chkUseNicks.TabIndex = 1;
            chkUseNicks.Text = "Use Nicknames";
            chkUseNicks.UseVisualStyleBackColor = true;
            // 
            // groupBox5
            // 
            groupBox5.Location = new System.Drawing.Point(12, 14);
            groupBox5.Name = "groupBox5";
            groupBox5.Size = new System.Drawing.Size(140, 100);
            groupBox5.TabIndex = 0;
            groupBox5.TabStop = false;
            groupBox5.Text = "Text Display";
            // 
            // tabPage3
            // 
            tabPage3.Controls.Add(label1);
            tabPage3.Controls.Add(txtTSFormat);
            tabPage3.Controls.Add(chkTimestamp);
            tabPage3.Controls.Add(chkParseUrls);
            tabPage3.Controls.Add(groupBox9);
            tabPage3.Location = new System.Drawing.Point(4, 24);
            tabPage3.Name = "tabPage3";
            tabPage3.Size = new System.Drawing.Size(332, 315);
            tabPage3.TabIndex = 2;
            tabPage3.Text = "Chat Windows";
            tabPage3.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            label1.Location = new System.Drawing.Point(205, 56);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(63, 15);
            label1.TabIndex = 4;
            label1.Text = "Format";
            // 
            // txtTSFormat
            // 
            txtTSFormat.Location = new System.Drawing.Point(205, 75);
            txtTSFormat.Name = "txtTSFormat";
            txtTSFormat.PlaceholderText = "[hh:mm:ss]";
            txtTSFormat.Size = new System.Drawing.Size(96, 23);
            txtTSFormat.TabIndex = 3;
            // 
            // chkTimestamp
            // 
            chkTimestamp.Location = new System.Drawing.Point(205, 31);
            chkTimestamp.Name = "chkTimestamp";
            chkTimestamp.Size = new System.Drawing.Size(99, 22);
            chkTimestamp.TabIndex = 2;
            chkTimestamp.Text = "Timestamp";
            chkTimestamp.UseVisualStyleBackColor = true;
            // 
            // chkParseUrls
            // 
            chkParseUrls.Location = new System.Drawing.Point(205, 3);
            chkParseUrls.Name = "chkParseUrls";
            chkParseUrls.Size = new System.Drawing.Size(99, 22);
            chkParseUrls.TabIndex = 1;
            chkParseUrls.Text = "Parse URLs";
            chkParseUrls.UseVisualStyleBackColor = true;
            // 
            // groupBox9
            // 
            groupBox9.Controls.Add(chkShowNickChange);
            groupBox9.Controls.Add(chkReturnFromBattle);
            groupBox9.Controls.Add(chkShowTeamChange);
            groupBox9.Controls.Add(chkBattleResults);
            groupBox9.Controls.Add(chkBattleStart);
            groupBox9.Controls.Add(chkShowIgnore);
            groupBox9.Controls.Add(chkPlayerSignOff);
            groupBox9.Controls.Add(chkPlayerSignOn);
            groupBox9.Controls.Add(chkShowBack);
            groupBox9.Controls.Add(chkShowAway);
            groupBox9.Location = new System.Drawing.Point(8, 3);
            groupBox9.Name = "groupBox9";
            groupBox9.Size = new System.Drawing.Size(191, 276);
            groupBox9.TabIndex = 0;
            groupBox9.TabStop = false;
            groupBox9.Text = "Message Toggles";
            // 
            // chkShowNickChange
            // 
            chkShowNickChange.Location = new System.Drawing.Point(6, 247);
            chkShowNickChange.Name = "chkShowNickChange";
            chkShowNickChange.Size = new System.Drawing.Size(166, 19);
            chkShowNickChange.TabIndex = 9;
            chkShowNickChange.Text = "Show Nickname Change";
            chkShowNickChange.UseVisualStyleBackColor = true;
            // 
            // chkReturnFromBattle
            // 
            chkReturnFromBattle.Location = new System.Drawing.Point(6, 122);
            chkReturnFromBattle.Name = "chkReturnFromBattle";
            chkReturnFromBattle.Size = new System.Drawing.Size(166, 19);
            chkReturnFromBattle.TabIndex = 4;
            chkReturnFromBattle.Text = "Show Return from Battle";
            chkReturnFromBattle.UseVisualStyleBackColor = true;
            // 
            // chkShowTeamChange
            // 
            chkShowTeamChange.Location = new System.Drawing.Point(6, 222);
            chkShowTeamChange.Name = "chkShowTeamChange";
            chkShowTeamChange.Size = new System.Drawing.Size(166, 19);
            chkShowTeamChange.TabIndex = 8;
            chkShowTeamChange.Text = "Show Team Change";
            chkShowTeamChange.UseVisualStyleBackColor = true;
            // 
            // chkBattleResults
            // 
            chkBattleResults.Location = new System.Drawing.Point(6, 97);
            chkBattleResults.Name = "chkBattleResults";
            chkBattleResults.Size = new System.Drawing.Size(166, 19);
            chkBattleResults.TabIndex = 3;
            chkBattleResults.Text = "Show Battle Results";
            chkBattleResults.UseVisualStyleBackColor = true;
            // 
            // chkBattleStart
            // 
            chkBattleStart.Location = new System.Drawing.Point(6, 72);
            chkBattleStart.Name = "chkBattleStart";
            chkBattleStart.Size = new System.Drawing.Size(166, 19);
            chkBattleStart.TabIndex = 2;
            chkBattleStart.Text = "Show Battle Start";
            chkBattleStart.UseVisualStyleBackColor = true;
            // 
            // chkShowIgnore
            // 
            chkShowIgnore.Location = new System.Drawing.Point(6, 197);
            chkShowIgnore.Name = "chkShowIgnore";
            chkShowIgnore.Size = new System.Drawing.Size(166, 19);
            chkShowIgnore.TabIndex = 7;
            chkShowIgnore.Text = "Show Ignore/Unignore";
            chkShowIgnore.UseVisualStyleBackColor = true;
            // 
            // chkPlayerSignOff
            // 
            chkPlayerSignOff.Location = new System.Drawing.Point(6, 47);
            chkPlayerSignOff.Name = "chkPlayerSignOff";
            chkPlayerSignOff.Size = new System.Drawing.Size(166, 19);
            chkPlayerSignOff.TabIndex = 1;
            chkPlayerSignOff.Text = "Show Player Sign Off";
            chkPlayerSignOff.UseVisualStyleBackColor = true;
            // 
            // chkPlayerSignOn
            // 
            chkPlayerSignOn.Location = new System.Drawing.Point(6, 22);
            chkPlayerSignOn.Name = "chkPlayerSignOn";
            chkPlayerSignOn.Size = new System.Drawing.Size(166, 19);
            chkPlayerSignOn.TabIndex = 0;
            chkPlayerSignOn.Text = "Show Player Sign On";
            chkPlayerSignOn.UseVisualStyleBackColor = true;
            // 
            // chkShowBack
            // 
            chkShowBack.Location = new System.Drawing.Point(6, 172);
            chkShowBack.Name = "chkShowBack";
            chkShowBack.Size = new System.Drawing.Size(166, 19);
            chkShowBack.TabIndex = 6;
            chkShowBack.Text = "Show back from away";
            chkShowBack.UseVisualStyleBackColor = true;
            // 
            // chkShowAway
            // 
            chkShowAway.Location = new System.Drawing.Point(6, 147);
            chkShowAway.Name = "chkShowAway";
            chkShowAway.Size = new System.Drawing.Size(166, 19);
            chkShowAway.TabIndex = 5;
            chkShowAway.Text = "Show Away";
            chkShowAway.UseVisualStyleBackColor = true;
            // 
            // btnOk
            // 
            btnOk.Location = new System.Drawing.Point(158, 345);
            btnOk.Name = "btnOk";
            btnOk.Size = new System.Drawing.Size(82, 32);
            btnOk.TabIndex = 2;
            btnOk.Text = "&Ok";
            btnOk.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            btnCancel.Location = new System.Drawing.Point(246, 345);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new System.Drawing.Size(82, 32);
            btnCancel.TabIndex = 3;
            btnCancel.Text = "&Cancel";
            btnCancel.UseVisualStyleBackColor = true;
            // 
            // Settings
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(340, 382);
            Controls.Add(btnCancel);
            Controls.Add(btnOk);
            Controls.Add(tabControl1);
            Icon = ((System.Drawing.Icon)resources.GetObject("$this.Icon"));
            Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            MaximizeBox = false;
            MinimizeBox = false;
            Text = "Settings";
            Load += Settings_Load;
            tabControl1.ResumeLayout(false);
            tabPage1.ResumeLayout(false);
            groupBox3.ResumeLayout(false);
            groupBox4.ResumeLayout(false);
            groupBox4.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox1.ResumeLayout(false);
            tabPage2.ResumeLayout(false);
            groupBox8.ResumeLayout(false);
            groupBox7.ResumeLayout(false);
            groupBox6.ResumeLayout(false);
            tabPage3.ResumeLayout(false);
            tabPage3.PerformLayout();
            groupBox9.ResumeLayout(false);
            ResumeLayout(false);
        }

        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;

        private System.Windows.Forms.CheckBox chkPlayerSignOff;
        private System.Windows.Forms.CheckBox chkBattleStart;
        private System.Windows.Forms.CheckBox chkBattleResults;
        private System.Windows.Forms.CheckBox chkReturnFromBattle;
        private System.Windows.Forms.CheckBox chkShowNickChange;
        private System.Windows.Forms.CheckBox chkShowTeamChange;
        private System.Windows.Forms.CheckBox chkShowIgnore;
        private System.Windows.Forms.CheckBox chkShowBack;
        private System.Windows.Forms.CheckBox chkShowAway;

        private System.Windows.Forms.CheckBox chkParseUrls;
        private System.Windows.Forms.CheckBox chkTimestamp;
        private System.Windows.Forms.TextBox txtTSFormat;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkPlayerSignOn;

        private System.Windows.Forms.GroupBox groupBox9;

        private System.Windows.Forms.CheckBox chkAnimateHP;
        private System.Windows.Forms.CheckBox chkEnableBackground;

        private System.Windows.Forms.RadioButton rdoExtended;

        private System.Windows.Forms.CheckBox chkUseNicks;
        private System.Windows.Forms.CheckBox chkTrainerPrefix;
        private System.Windows.Forms.RadioButton rdoNormal;

        private System.Windows.Forms.GroupBox groupBox8;
        private System.Windows.Forms.CheckBox chkSounds;
        private System.Windows.Forms.CheckBox chkMusic;

        private System.Windows.Forms.GroupBox groupBox7;

        private System.Windows.Forms.GroupBox groupBox6;

        private System.Windows.Forms.GroupBox groupBox5;

        private System.Windows.Forms.Button btnStoredPassword;

        private System.Windows.Forms.CheckBox chkAutoLoad;
        private System.Windows.Forms.RadioButton rdoShowTeam;
        private System.Windows.Forms.RadioButton rdoHideTeam;

        private System.Windows.Forms.CheckBox chkReplayAutosave;
        private System.Windows.Forms.CheckBox chkReplays;

        private System.Windows.Forms.CheckBox chkLogsSave;
        private System.Windows.Forms.CheckBox chkLogsAutosave;

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;

        private System.Windows.Forms.TabPage tabPage3;

        private System.Windows.Forms.TextBox txtRegAddr;

        private System.Windows.Forms.Label lblRegAddr;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;

        #endregion
    }
}