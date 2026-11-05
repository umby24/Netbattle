namespace Netbattle.Forms
{
    partial class TrayIcon
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TrayIcon));
            this.pichook = new System.Windows.Forms.PictureBox();
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.mnuBar = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuMain_0 = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuMain_1 = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuMain_2 = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuMain_3 = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuMain_4 = new System.Windows.Forms.ToolStripMenuItem();
            this.Timer1 = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.pichook)).BeginInit();
            this.menuStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // pichook
            // 
            this.pichook.Location = new System.Drawing.Point(0, 24);
            this.pichook.Name = "pichook";
            this.pichook.Size = new System.Drawing.Size(57, 37);
            this.pichook.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pichook.TabStop = false;
            // 
            // menuStrip
            // 
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
                        this.mnuBar});
            this.menuStrip.Location = new System.Drawing.Point(0, 0);
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Size = new System.Drawing.Size(149, 24);
            this.menuStrip.TabIndex = 0;
            this.menuStrip.Text = "menuStrip";
            // 
            // mnuBar
            // 
            this.mnuBar.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
                        this.mnuMain_0,
                        this.mnuMain_1,
                        this.mnuMain_2,
                        this.mnuMain_3,
                        this.mnuMain_4});
            this.mnuBar.Name = "mnuBar";
            this.mnuBar.Text = "PopupMenu";
            // 
            // mnuMain_0
            // 
            this.mnuMain_0.Name = "mnuMain_0";
            this.mnuMain_0.Text = "&Restore Window";
            // 
            // mnuMain_1
            // 
            this.mnuMain_1.Name = "mnuMain_1";
            this.mnuMain_1.Text = "S&cript Window";
            // 
            // mnuMain_2
            // 
            this.mnuMain_2.Name = "mnuMain_2";
            this.mnuMain_2.Text = "&Set Options";
            // 
            // mnuMain_3
            // 
            this.mnuMain_3.Name = "mnuMain_3";
            this.mnuMain_3.Text = "&Data Manager";
            // 
            // mnuMain_4
            // 
            this.mnuMain_4.Name = "mnuMain_4";
            this.mnuMain_4.Text = "E&xit";
            // 
            // Timer1
            // 
            this.Timer1.Enabled = true;
            this.Timer1.Interval = 1000;
            // 
            // TrayIcon
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(149, 65);
            this.Controls.Add(this.pichook);
            this.Controls.Add(this.menuStrip);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(101, 304);
            this.MainMenuStrip = this.menuStrip;
            this.Name = "TrayIcon";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "TrayIcon";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pichook)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.PictureBox pichook;
        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStripMenuItem mnuBar;
        private System.Windows.Forms.ToolStripMenuItem mnuMain_0;
        private System.Windows.Forms.ToolStripMenuItem mnuMain_1;
        private System.Windows.Forms.ToolStripMenuItem mnuMain_2;
        private System.Windows.Forms.ToolStripMenuItem mnuMain_3;
        private System.Windows.Forms.ToolStripMenuItem mnuMain_4;
        private System.Windows.Forms.Timer Timer1;
    }
}
