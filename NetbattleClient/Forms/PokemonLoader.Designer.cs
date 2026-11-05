namespace Netbattle.Forms
{
    partial class PokemonLoader
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PokemonLoader));
            this.Progress = new System.Windows.Forms.ProgressBar();
            this.Label1 = new System.Windows.Forms.Label();
            this.Image1 = new System.Windows.Forms.PictureBox();
            this.AnimTimer = new System.Windows.Forms.Timer(this.components);
            this.ImageList1 = new System.Windows.Forms.ImageList(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.Image1)).BeginInit();
            this.SuspendLayout();
            // 
            // Progress
            // 
            this.Progress.Location = new System.Drawing.Point(8, 40);
            this.Progress.Name = "Progress";
            this.Progress.Size = new System.Drawing.Size(281, 25);
            // 
            // Label1
            // 
            this.Label1.Location = new System.Drawing.Point(16, 16);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(273, 17);
            this.Label1.TabIndex = 0;
            this.Label1.BackColor = System.Drawing.Color.Transparent;
            this.Label1.Text = "Loading Pokémon Data...";
            // 
            // Image1
            // 
            this.Image1.Location = new System.Drawing.Point(128, 72);
            this.Image1.Name = "Image1";
            this.Image1.Size = new System.Drawing.Size(32, 32);
            this.Image1.BackColor = System.Drawing.Color.Transparent;
            this.Image1.TabStop = false;
            // 
            // ImageList1
            // 
            this.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit;
            this.ImageList1.ImageSize = new System.Drawing.Size(32, 32);
            this.ImageList1.TransparentColor = System.Drawing.Color.FromArgb(192, 192, 192);
            this.ImageList1.Images.Add(((System.Drawing.Image)(resources.GetObject("ImageList1.Image1"))));
            this.ImageList1.Images.Add(((System.Drawing.Image)(resources.GetObject("ImageList1.Image2"))));
            this.ImageList1.Images.Add(((System.Drawing.Image)(resources.GetObject("ImageList1.Image3"))));
            // 
            // AnimTimer
            // 
            this.AnimTimer.Enabled = true;
            this.AnimTimer.Interval = 250;
            // 
            // PokemonLoader
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(297, 113);
            this.ControlBox = false;
            this.Controls.Add(this.Progress);
            this.Controls.Add(this.Label1);
            this.Controls.Add(this.Image1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Location = new System.Drawing.Point(3, 3);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "PokemonLoader";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "";
            ((System.ComponentModel.ISupportInitialize)(this.Image1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.ProgressBar Progress;
        private System.Windows.Forms.Label Label1;
        private System.Windows.Forms.PictureBox Image1;
        private System.Windows.Forms.Timer AnimTimer;
        private System.Windows.Forms.ImageList ImageList1;
    }
}
