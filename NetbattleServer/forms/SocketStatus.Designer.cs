namespace Netbattle.Forms
{
    partial class SocketStatus
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SocketStatus));
            this.Command2 = new System.Windows.Forms.Button();
            this.Command1 = new System.Windows.Forms.Button();
            this.SocketList = new System.Windows.Forms.ListView();
            this.SocketList_Col1 = new System.Windows.Forms.ColumnHeader();
            this.SocketList_Col2 = new System.Windows.Forms.ColumnHeader();
            this.SocketList_Col3 = new System.Windows.Forms.ColumnHeader();
            this.Label1 = new System.Windows.Forms.Label();
            this.RefreshTimer = new System.Windows.Forms.Timer(this.components);
            this.ImageList1 = new System.Windows.Forms.ImageList(this.components);
            this.SuspendLayout();
            // 
            // Command2
            // 
            this.Command2.Location = new System.Drawing.Point(240, 312);
            this.Command2.Name = "Command2";
            this.Command2.Size = new System.Drawing.Size(81, 33);
            this.Command2.TabIndex = 3;
            this.Command2.Text = "&Close";
            this.Command2.UseVisualStyleBackColor = true;
            // 
            // Command1
            // 
            this.Command1.Location = new System.Drawing.Point(152, 312);
            this.Command1.Name = "Command1";
            this.Command1.Size = new System.Drawing.Size(81, 33);
            this.Command1.TabIndex = 2;
            this.Command1.Text = "&Disconnect";
            this.Command1.UseVisualStyleBackColor = true;
            // 
            // SocketList
            // 
            this.SocketList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
                        this.SocketList_Col1,
                        this.SocketList_Col2,
                        this.SocketList_Col3});
            this.SocketList.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.SocketList.LabelEdit = true;
            this.SocketList.MultiSelect = false;
            this.SocketList.SmallImageList = this.ImageList1;
            this.SocketList.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.SocketList.Location = new System.Drawing.Point(8, 24);
            this.SocketList.Name = "SocketList";
            this.SocketList.Size = new System.Drawing.Size(313, 281);
            this.SocketList.TabIndex = 0;
            this.SocketList.UseCompatibleStateImageBehavior = false;
            this.SocketList.View = System.Windows.Forms.View.Details;
            // 
            // SocketList_Col1
            // 
            this.SocketList_Col1.Text = "Number";
            this.SocketList_Col1.Width = 169;
            // 
            // SocketList_Col2
            // 
            this.SocketList_Col2.Text = "Player";
            this.SocketList_Col2.Width = 169;
            // 
            // SocketList_Col3
            // 
            this.SocketList_Col3.Text = "State";
            this.SocketList_Col3.Width = 169;
            // 
            // Label1
            // 
            this.Label1.Location = new System.Drawing.Point(8, 8);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(313, 17);
            this.Label1.TabIndex = 1;
            this.Label1.BackColor = System.Drawing.Color.Transparent;
            this.Label1.Text = "Available Sockets";
            // 
            // ImageList1
            // 
            this.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit;
            this.ImageList1.ImageSize = new System.Drawing.Size(16, 16);
            this.ImageList1.TransparentColor = System.Drawing.Color.FromArgb(192, 192, 192);
            this.ImageList1.Images.Add(((System.Drawing.Image)(resources.GetObject("ImageList1.Image1"))));
            this.ImageList1.Images.Add(((System.Drawing.Image)(resources.GetObject("ImageList1.Image2"))));
            this.ImageList1.Images.Add(((System.Drawing.Image)(resources.GetObject("ImageList1.Image3"))));
            this.ImageList1.Images.Add(((System.Drawing.Image)(resources.GetObject("ImageList1.Image4"))));
            // 
            // RefreshTimer
            // 
            this.RefreshTimer.Enabled = true;
            this.RefreshTimer.Interval = 1000;
            // 
            // SocketStatus
            // 
            this.AcceptButton = this.Command2;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.Command2;
            this.ClientSize = new System.Drawing.Size(330, 351);
            this.Controls.Add(this.Command2);
            this.Controls.Add(this.Command1);
            this.Controls.Add(this.SocketList);
            this.Controls.Add(this.Label1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(3, 29);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SocketStatus";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Socket Status";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Button Command2;
        private System.Windows.Forms.Button Command1;
        private System.Windows.Forms.ListView SocketList;
        private System.Windows.Forms.ColumnHeader SocketList_Col1;
        private System.Windows.Forms.ColumnHeader SocketList_Col2;
        private System.Windows.Forms.ColumnHeader SocketList_Col3;
        private System.Windows.Forms.Label Label1;
        private System.Windows.Forms.Timer RefreshTimer;
        private System.Windows.Forms.ImageList ImageList1;
    }
}
