namespace Netbattle.Forms
{
    partial class Splash
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Splash));
            this.Label1 = new System.Windows.Forms.Label();
            this.Image1 = new System.Windows.Forms.Panel();
            this.lblVersion = new System.Windows.Forms.Label();
            this.Image1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Label1
            // 
            this.Label1.Location = new System.Drawing.Point(256, 144);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(129, 17);
            this.Label1.TabIndex = 1;
            this.Label1.Text = "Scripters NB V 1";
            // 
            // Image1
            // 
            this.Image1.Controls.Add(this.lblVersion);
            this.Image1.Location = new System.Drawing.Point(0, 0);
            this.Image1.Name = "Image1";
            this.Image1.Size = new System.Drawing.Size(400, 250);
            this.Image1.BackColor = System.Drawing.Color.Transparent;
            this.Image1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("Image1.BackgroundImage")));
            this.Image1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            // 
            // lblVersion
            // 
            this.lblVersion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVersion.Location = new System.Drawing.Point(312, 223);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(81, 17);
            this.lblVersion.TabIndex = 0;
            this.lblVersion.BackColor = System.Drawing.Color.Transparent;
            this.lblVersion.Text = "v0.0.0";
            this.lblVersion.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // Splash
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(400, 250);
            this.ControlBox = false;
            this.Controls.Add(this.Label1);
            this.Controls.Add(this.Image1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Splash";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "";
            this.Image1.ResumeLayout(false);
            this.Image1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Label Label1;
        private System.Windows.Forms.Panel Image1;
        private System.Windows.Forms.Label lblVersion;
    }
}
