using System.ComponentModel;

namespace Netbattle.Forms
{
    partial class About
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
            ComponentResourceManager resources = new ComponentResourceManager(typeof(About));
            btnOk = new Button();
            btnSystemInfo = new Button();
            label1 = new Label();
            label2 = new Label();
            linkLabel1 = new LinkLabel();
            SuspendLayout();
            // 
            // btnOk
            // 
            btnOk.Location = new Point(555, 222);
            btnOk.Margin = new Padding(4, 3, 4, 3);
            btnOk.Name = "btnOk";
            btnOk.Size = new Size(97, 33);
            btnOk.TabIndex = 0;
            btnOk.Text = "OK";
            btnOk.UseVisualStyleBackColor = true;
            btnOk.Click += btnOk_Click;
            // 
            // btnSystemInfo
            // 
            btnSystemInfo.Location = new Point(450, 222);
            btnSystemInfo.Margin = new Padding(4, 3, 4, 3);
            btnSystemInfo.Name = "btnSystemInfo";
            btnSystemInfo.Size = new Size(97, 33);
            btnSystemInfo.TabIndex = 1;
            btnSystemInfo.Text = "System Info...";
            btnSystemInfo.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            label1.Font = new Font("Microsoft Sans Serif", 26.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label1.Location = new Point(13, 9);
            label1.Margin = new Padding(4, 0, 4, 0);
            label1.Name = "label1";
            label1.Size = new Size(424, 51);
            label1.TabIndex = 2;
            label1.Text = "Pokemon Netbattle (C#)";
            // 
            // label2
            // 
            label2.Location = new Point(16, 74);
            label2.Margin = new Padding(4, 0, 4, 0);
            label2.Name = "label2";
            label2.Size = new Size(644, 239);
            label2.TabIndex = 3;
            label2.Text = resources.GetString("label2.Text");
            // 
            // linkLabel1
            // 
            linkLabel1.Location = new Point(210, 182);
            linkLabel1.Margin = new Padding(4, 0, 4, 0);
            linkLabel1.Name = "linkLabel1";
            linkLabel1.Size = new Size(227, 20);
            linkLabel1.TabIndex = 4;
            linkLabel1.TabStop = true;
            linkLabel1.Text = "https://github.com/umby24/Netbattle";
            // 
            // About
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(683, 280);
            ControlBox = false;
            Controls.Add(linkLabel1);
            Controls.Add(label1);
            Controls.Add(btnSystemInfo);
            Controls.Add(btnOk);
            Controls.Add(label2);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Margin = new Padding(4, 3, 4, 3);
            Name = "About";
            Text = "About";
            ResumeLayout(false);
        }

        private System.Windows.Forms.LinkLabel linkLabel1;

        private System.Windows.Forms.Label label2;

        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnSystemInfo;
        private System.Windows.Forms.Label label1;

        #endregion
    }
}