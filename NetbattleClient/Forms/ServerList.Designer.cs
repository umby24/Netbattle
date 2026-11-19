namespace NetbattleClient.Forms {
    partial class ServerList {
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
            lstServers = new System.Windows.Forms.ListView();
            colName = new System.Windows.Forms.ColumnHeader();
            colAdmin = new System.Windows.Forms.ColumnHeader();
            colUsers = new System.Windows.Forms.ColumnHeader();
            btnCancel = new System.Windows.Forms.Button();
            btnConnect = new System.Windows.Forms.Button();
            btnDirect = new System.Windows.Forms.Button();
            txtDescription = new System.Windows.Forms.TextBox();
            statusStrip1 = new System.Windows.Forms.StatusStrip();
            lblStatus = new System.Windows.Forms.ToolStripStatusLabel();
            lblChoose = new System.Windows.Forms.Label();
            statusStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // lstServers
            // 
            lstServers.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] { colName, colAdmin, colUsers });
            lstServers.GridLines = true;
            lstServers.Location = new System.Drawing.Point(0, 35);
            lstServers.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            lstServers.Name = "lstServers";
            lstServers.Size = new System.Drawing.Size(331, 192);
            lstServers.TabIndex = 0;
            lstServers.UseCompatibleStateImageBehavior = false;
            lstServers.View = System.Windows.Forms.View.Details;
            lstServers.SelectedIndexChanged += lstServers_SelectedIndexChanged;
            lstServers.MouseDoubleClick += lstServers_MouseDoubleClick;
            // 
            // colName
            // 
            colName.Name = "colName";
            colName.Text = "Name";
            colName.Width = 138;
            // 
            // colAdmin
            // 
            colAdmin.Name = "colAdmin";
            colAdmin.Text = "Main Admin";
            colAdmin.Width = 118;
            // 
            // colUsers
            // 
            colUsers.Name = "colUsers";
            colUsers.Text = "Users/Max";
            colUsers.Width = 70;
            // 
            // btnCancel
            // 
            btnCancel.Location = new System.Drawing.Point(244, 293);
            btnCancel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new System.Drawing.Size(87, 27);
            btnCancel.TabIndex = 1;
            btnCancel.Text = "Cancel";
            btnCancel.UseVisualStyleBackColor = true;
            btnCancel.Click += btnCancel_Click;
            // 
            // btnConnect
            // 
            btnConnect.Location = new System.Drawing.Point(122, 293);
            btnConnect.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btnConnect.Name = "btnConnect";
            btnConnect.Size = new System.Drawing.Size(114, 27);
            btnConnect.TabIndex = 2;
            btnConnect.Text = "Connect";
            btnConnect.UseVisualStyleBackColor = true;
            btnConnect.Click += btnConnect_Click;
            // 
            // btnDirect
            // 
            btnDirect.Location = new System.Drawing.Point(0, 293);
            btnDirect.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btnDirect.Name = "btnDirect";
            btnDirect.Size = new System.Drawing.Size(114, 27);
            btnDirect.TabIndex = 3;
            btnDirect.Text = "Advanced";
            btnDirect.UseVisualStyleBackColor = true;
            btnDirect.Click += btnDirect_Click;
            // 
            // txtDescription
            // 
            txtDescription.Location = new System.Drawing.Point(0, 233);
            txtDescription.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            txtDescription.Multiline = true;
            txtDescription.Name = "txtDescription";
            txtDescription.Size = new System.Drawing.Size(331, 54);
            txtDescription.TabIndex = 4;
            // 
            // statusStrip1
            // 
            statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { lblStatus });
            statusStrip1.Location = new System.Drawing.Point(0, 323);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new System.Drawing.Size(331, 22);
            statusStrip1.TabIndex = 5;
            statusStrip1.Text = "statusStrip1";
            // 
            // lblStatus
            // 
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new System.Drawing.Size(79, 17);
            lblStatus.Text = "Disconnected";
            // 
            // lblChoose
            // 
            lblChoose.Font = new System.Drawing.Font("Verdana", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)0));
            lblChoose.Location = new System.Drawing.Point(30, 1);
            lblChoose.Name = "lblChoose";
            lblChoose.Size = new System.Drawing.Size(289, 32);
            lblChoose.TabIndex = 6;
            lblChoose.Text = "Please choose a server.";
            // 
            // ServerList
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(331, 345);
            Controls.Add(lblChoose);
            Controls.Add(statusStrip1);
            Controls.Add(txtDescription);
            Controls.Add(btnDirect);
            Controls.Add(btnConnect);
            Controls.Add(btnCancel);
            Controls.Add(lstServers);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Server Listing";
            FormClosed += ServerList_FormClosed;
            Load += ServerList_Load;
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        private System.Windows.Forms.Label lblChoose;

        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel lblStatus;

        #endregion

        private System.Windows.Forms.ListView lstServers;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnConnect;
        private System.Windows.Forms.Button btnDirect;
        private System.Windows.Forms.TextBox txtDescription;
        private System.Windows.Forms.ColumnHeader colName;
        private System.Windows.Forms.ColumnHeader colAdmin;
        private System.Windows.Forms.ColumnHeader colUsers;
    }
}