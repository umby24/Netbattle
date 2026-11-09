using System.ComponentModel;

namespace Netbattle.Forms
{
    partial class MasterServer
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
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MasterServer));
            this.mainMenu1 = new System.Windows.Forms.MainMenu(this.components);
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveLogToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.startLoggingToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.serverToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.serverDataToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.socketStatusToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.scriptWindowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.databaseChangesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripSeparator();
            this.replySpyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.rawDataToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.confirmExitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripSeparator();
            this.clearMessagesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.reloadDatabaseFilesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.playerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.infoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.kickToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.banToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.banIDNameToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripSeparator();
            this.refreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.windowsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lstUsers = new System.Windows.Forms.ListBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.txtChatBox = new System.Windows.Forms.TextBox();
            this.btnSend = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { this.fileToolStripMenuItem, this.serverToolStripMenuItem, this.playerToolStripMenuItem, this.windowsToolStripMenuItem, this.helpToolStripMenuItem });
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(440, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { this.saveLogToolStripMenuItem, this.startLoggingToolStripMenuItem, this.toolStripMenuItem1, this.exitToolStripMenuItem });
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "&File";
            // 
            // saveLogToolStripMenuItem
            // 
            this.saveLogToolStripMenuItem.Name = "saveLogToolStripMenuItem";
            this.saveLogToolStripMenuItem.Size = new System.Drawing.Size(145, 22);
            this.saveLogToolStripMenuItem.Text = "Save Log...";
            // 
            // startLoggingToolStripMenuItem
            // 
            this.startLoggingToolStripMenuItem.Name = "startLoggingToolStripMenuItem";
            this.startLoggingToolStripMenuItem.Size = new System.Drawing.Size(145, 22);
            this.startLoggingToolStripMenuItem.Text = "Start Logging";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(142, 6);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(145, 22);
            this.exitToolStripMenuItem.Text = "E&xit";
            // 
            // serverToolStripMenuItem
            // 
            this.serverToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { this.optionsToolStripMenuItem, this.serverDataToolStripMenuItem, this.socketStatusToolStripMenuItem, this.scriptWindowToolStripMenuItem, this.databaseChangesToolStripMenuItem, this.toolStripMenuItem2, this.replySpyToolStripMenuItem, this.rawDataToolStripMenuItem, this.confirmExitToolStripMenuItem, this.toolStripMenuItem3, this.clearMessagesToolStripMenuItem, this.reloadDatabaseFilesToolStripMenuItem });
            this.serverToolStripMenuItem.Name = "serverToolStripMenuItem";
            this.serverToolStripMenuItem.Size = new System.Drawing.Size(51, 20);
            this.serverToolStripMenuItem.Text = "&Server";
            // 
            // optionsToolStripMenuItem
            // 
            this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            this.optionsToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.optionsToolStripMenuItem.Text = "Options...";
            // 
            // serverDataToolStripMenuItem
            // 
            this.serverDataToolStripMenuItem.Name = "serverDataToolStripMenuItem";
            this.serverDataToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.serverDataToolStripMenuItem.Text = "Server Data...";
            // 
            // socketStatusToolStripMenuItem
            // 
            this.socketStatusToolStripMenuItem.Name = "socketStatusToolStripMenuItem";
            this.socketStatusToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.socketStatusToolStripMenuItem.Text = "Socket Status...";
            // 
            // scriptWindowToolStripMenuItem
            // 
            this.scriptWindowToolStripMenuItem.Name = "scriptWindowToolStripMenuItem";
            this.scriptWindowToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.scriptWindowToolStripMenuItem.Text = "Script Window...";
            // 
            // databaseChangesToolStripMenuItem
            // 
            this.databaseChangesToolStripMenuItem.Name = "databaseChangesToolStripMenuItem";
            this.databaseChangesToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.databaseChangesToolStripMenuItem.Text = "Database Changes...";
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(184, 6);
            // 
            // replySpyToolStripMenuItem
            // 
            this.replySpyToolStripMenuItem.Name = "replySpyToolStripMenuItem";
            this.replySpyToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.replySpyToolStripMenuItem.Text = "Reply Spy";
            // 
            // rawDataToolStripMenuItem
            // 
            this.rawDataToolStripMenuItem.Name = "rawDataToolStripMenuItem";
            this.rawDataToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.rawDataToolStripMenuItem.Text = "Raw Data";
            // 
            // confirmExitToolStripMenuItem
            // 
            this.confirmExitToolStripMenuItem.Checked = true;
            this.confirmExitToolStripMenuItem.CheckOnClick = true;
            this.confirmExitToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.confirmExitToolStripMenuItem.Name = "confirmExitToolStripMenuItem";
            this.confirmExitToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.confirmExitToolStripMenuItem.Text = "Confirm Exit";
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(184, 6);
            // 
            // clearMessagesToolStripMenuItem
            // 
            this.clearMessagesToolStripMenuItem.Name = "clearMessagesToolStripMenuItem";
            this.clearMessagesToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.clearMessagesToolStripMenuItem.Text = "Clear Messages";
            // 
            // reloadDatabaseFilesToolStripMenuItem
            // 
            this.reloadDatabaseFilesToolStripMenuItem.Name = "reloadDatabaseFilesToolStripMenuItem";
            this.reloadDatabaseFilesToolStripMenuItem.Size = new System.Drawing.Size(187, 22);
            this.reloadDatabaseFilesToolStripMenuItem.Text = "Reload Database Files";
            // 
            // playerToolStripMenuItem
            // 
            this.playerToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { this.infoToolStripMenuItem, this.kickToolStripMenuItem, this.banToolStripMenuItem, this.banIDNameToolStripMenuItem, this.toolStripMenuItem4, this.refreshToolStripMenuItem });
            this.playerToolStripMenuItem.Name = "playerToolStripMenuItem";
            this.playerToolStripMenuItem.Size = new System.Drawing.Size(51, 20);
            this.playerToolStripMenuItem.Text = "&Player";
            // 
            // infoToolStripMenuItem
            // 
            this.infoToolStripMenuItem.Name = "infoToolStripMenuItem";
            this.infoToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.infoToolStripMenuItem.Text = "Info...";
            // 
            // kickToolStripMenuItem
            // 
            this.kickToolStripMenuItem.Name = "kickToolStripMenuItem";
            this.kickToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.kickToolStripMenuItem.Text = "Kick";
            // 
            // banToolStripMenuItem
            // 
            this.banToolStripMenuItem.Name = "banToolStripMenuItem";
            this.banToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.banToolStripMenuItem.Text = "Ban (IP)";
            this.banToolStripMenuItem.Click += new System.EventHandler(this.banToolStripMenuItem_Click);
            // 
            // banIDNameToolStripMenuItem
            // 
            this.banIDNameToolStripMenuItem.Name = "banIDNameToolStripMenuItem";
            this.banIDNameToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.banIDNameToolStripMenuItem.Text = "Ban (ID/Name)";
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(150, 6);
            // 
            // refreshToolStripMenuItem
            // 
            this.refreshToolStripMenuItem.Name = "refreshToolStripMenuItem";
            this.refreshToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.refreshToolStripMenuItem.Text = "Refresh";
            // 
            // windowsToolStripMenuItem
            // 
            this.windowsToolStripMenuItem.Name = "windowsToolStripMenuItem";
            this.windowsToolStripMenuItem.Size = new System.Drawing.Size(63, 20);
            this.windowsToolStripMenuItem.Text = "&Window";
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "&Help";
            // 
            // lstUsers
            // 
            this.lstUsers.FormattingEnabled = true;
            this.lstUsers.Location = new System.Drawing.Point(12, 27);
            this.lstUsers.Name = "lstUsers";
            this.lstUsers.Size = new System.Drawing.Size(129, 238);
            this.lstUsers.TabIndex = 1;
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(147, 27);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(277, 191);
            this.richTextBox1.TabIndex = 2;
            this.richTextBox1.Text = "";
            // 
            // txtChatBox
            // 
            this.txtChatBox.Location = new System.Drawing.Point(147, 224);
            this.txtChatBox.Name = "txtChatBox";
            this.txtChatBox.Size = new System.Drawing.Size(277, 20);
            this.txtChatBox.TabIndex = 3;
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(349, 250);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(75, 23);
            this.btnSend.TabIndex = 4;
            this.btnSend.Text = "&Send";
            this.btnSend.UseVisualStyleBackColor = true;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Location = new System.Drawing.Point(0, 277);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(440, 22);
            this.statusStrip1.TabIndex = 5;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // MasterServer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(440, 299);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.txtChatBox);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.lstUsers);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Menu = this.mainMenu1;
            this.Name = "MasterServer";
            this.Text = "Netbattle Master Server";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private System.Windows.Forms.TextBox txtChatBox;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.StatusStrip statusStrip1;

        private System.Windows.Forms.RichTextBox richTextBox1;

        private System.Windows.Forms.ListBox lstUsers;

        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem4;
        private System.Windows.Forms.ToolStripMenuItem refreshToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem banIDNameToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem infoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem kickToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem banToolStripMenuItem;

        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem clearMessagesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem reloadDatabaseFilesToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem confirmExitToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem rawDataToolStripMenuItem;

        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem replySpyToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem databaseChangesToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem scriptWindowToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem serverDataToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem socketStatusToolStripMenuItem;

        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;

        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem saveLogToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem startLoggingToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;

        private System.Windows.Forms.MainMenu mainMenu1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem serverToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem playerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem windowsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;

        #endregion
    }
}