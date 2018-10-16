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
            this.components = new System.ComponentModel.Container();
            this.txtChat = new System.Windows.Forms.RichTextBox();
            this.lstPlayers = new System.Windows.Forms.ListView();
            this.imgColumn = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.nameColumn = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.txtInput = new System.Windows.Forms.TextBox();
            this.btnSend = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.playerMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.challengeInfoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.watchBattleToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.privateMessageToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.controlWindowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.kickToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.awayToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.playerMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtChat
            // 
            this.txtChat.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtChat.Location = new System.Drawing.Point(139, 32);
            this.txtChat.Name = "txtChat";
            this.txtChat.ReadOnly = true;
            this.txtChat.Size = new System.Drawing.Size(708, 340);
            this.txtChat.TabIndex = 0;
            this.txtChat.Text = "";
            // 
            // lstPlayers
            // 
            this.lstPlayers.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.imgColumn,
            this.nameColumn});
            this.lstPlayers.ContextMenuStrip = this.playerMenu;
            this.lstPlayers.FullRowSelect = true;
            this.lstPlayers.Location = new System.Drawing.Point(12, 32);
            this.lstPlayers.Name = "lstPlayers";
            this.lstPlayers.Size = new System.Drawing.Size(121, 373);
            this.lstPlayers.TabIndex = 1;
            this.lstPlayers.UseCompatibleStateImageBehavior = false;
            this.lstPlayers.View = System.Windows.Forms.View.Details;
            // 
            // imgColumn
            // 
            this.imgColumn.Width = 40;
            // 
            // nameColumn
            // 
            this.nameColumn.Text = "Name";
            this.nameColumn.Width = 75;
            // 
            // txtInput
            // 
            this.txtInput.Location = new System.Drawing.Point(139, 385);
            this.txtInput.MaxLength = 200;
            this.txtInput.Name = "txtInput";
            this.txtInput.Size = new System.Drawing.Size(627, 20);
            this.txtInput.TabIndex = 2;
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(772, 382);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(75, 23);
            this.btnSend.TabIndex = 3;
            this.btnSend.Text = "Send";
            this.btnSend.UseVisualStyleBackColor = true;
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.optionsToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(859, 24);
            this.menuStrip1.TabIndex = 4;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // playerMenu
            // 
            this.playerMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.challengeInfoToolStripMenuItem,
            this.watchBattleToolStripMenuItem,
            this.privateMessageToolStripMenuItem,
            this.controlWindowToolStripMenuItem,
            this.toolStripMenuItem1,
            this.kickToolStripMenuItem});
            this.playerMenu.Name = "playerMenu";
            this.playerMenu.Size = new System.Drawing.Size(162, 120);
            // 
            // challengeInfoToolStripMenuItem
            // 
            this.challengeInfoToolStripMenuItem.Name = "challengeInfoToolStripMenuItem";
            this.challengeInfoToolStripMenuItem.Size = new System.Drawing.Size(161, 22);
            this.challengeInfoToolStripMenuItem.Text = "Challenge/Info";
            // 
            // watchBattleToolStripMenuItem
            // 
            this.watchBattleToolStripMenuItem.Name = "watchBattleToolStripMenuItem";
            this.watchBattleToolStripMenuItem.Size = new System.Drawing.Size(161, 22);
            this.watchBattleToolStripMenuItem.Text = "Watch Battle";
            // 
            // privateMessageToolStripMenuItem
            // 
            this.privateMessageToolStripMenuItem.Name = "privateMessageToolStripMenuItem";
            this.privateMessageToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.privateMessageToolStripMenuItem.Text = "Private Message";
            this.privateMessageToolStripMenuItem.Click += new System.EventHandler(this.privateMessageToolStripMenuItem_Click);
            // 
            // controlWindowToolStripMenuItem
            // 
            this.controlWindowToolStripMenuItem.Name = "controlWindowToolStripMenuItem";
            this.controlWindowToolStripMenuItem.Size = new System.Drawing.Size(161, 22);
            this.controlWindowToolStripMenuItem.Text = "Control Window";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(158, 6);
            // 
            // kickToolStripMenuItem
            // 
            this.kickToolStripMenuItem.Name = "kickToolStripMenuItem";
            this.kickToolStripMenuItem.Size = new System.Drawing.Size(161, 22);
            this.kickToolStripMenuItem.Text = "&Kick";
            // 
            // optionsToolStripMenuItem
            // 
            this.optionsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.awayToolStripMenuItem});
            this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            this.optionsToolStripMenuItem.Size = new System.Drawing.Size(61, 20);
            this.optionsToolStripMenuItem.Text = "&Options";
            // 
            // awayToolStripMenuItem
            // 
            this.awayToolStripMenuItem.Name = "awayToolStripMenuItem";
            this.awayToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.awayToolStripMenuItem.Text = "&Away";
            // 
            // ServerWindow
            // 
            this.AcceptButton = this.btnSend;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(859, 417);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.txtInput);
            this.Controls.Add(this.lstPlayers);
            this.Controls.Add(this.txtChat);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "ServerWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Stadium:";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ServerWindow_FormClosing);
            this.Load += new System.EventHandler(this.ServerWindow_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.playerMenu.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox txtChat;
        private System.Windows.Forms.ListView lstPlayers;
        private System.Windows.Forms.TextBox txtInput;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.ColumnHeader imgColumn;
        private System.Windows.Forms.ColumnHeader nameColumn;
        private System.Windows.Forms.ContextMenuStrip playerMenu;
        private System.Windows.Forms.ToolStripMenuItem challengeInfoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem watchBattleToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem privateMessageToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem controlWindowToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem kickToolStripMenuItem;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem awayToolStripMenuItem;
    }
}