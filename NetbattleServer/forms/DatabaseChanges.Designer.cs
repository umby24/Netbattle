namespace Netbattle.Forms
{
    partial class DatabaseChanges
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
            this.Command2 = new System.Windows.Forms.Button();
            this.Command1 = new System.Windows.Forms.Button();
            this.txtDB = new System.Windows.Forms.TextBox();
            this.Label5 = new System.Windows.Forms.Label();
            this.Label4 = new System.Windows.Forms.Label();
            this.Label3 = new System.Windows.Forms.Label();
            this.Label2 = new System.Windows.Forms.Label();
            this.Label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Command2
            // 
            this.Command2.Location = new System.Drawing.Point(4, 256);
            this.Command2.Name = "Command2";
            this.Command2.Size = new System.Drawing.Size(129, 25);
            this.Command2.TabIndex = 7;
            this.Command2.Text = "OK";
            this.Command2.UseVisualStyleBackColor = true;
            // 
            // Command1
            // 
            this.Command1.Location = new System.Drawing.Point(4, 288);
            this.Command1.Name = "Command1";
            this.Command1.Size = new System.Drawing.Size(129, 25);
            this.Command1.TabIndex = 1;
            this.Command1.Text = "Cancel";
            this.Command1.UseVisualStyleBackColor = true;
            // 
            // txtDB
            // 
            this.txtDB.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDB.Location = new System.Drawing.Point(4, 8);
            this.txtDB.Name = "txtDB";
            this.txtDB.Size = new System.Drawing.Size(461, 157);
            this.txtDB.TabIndex = 0;
            this.txtDB.HideSelection = false;
            this.txtDB.Multiline = true;
            this.txtDB.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtDB.Text = "";
            // 
            // Label5
            // 
            this.Label5.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label5.Location = new System.Drawing.Point(152, 264);
            this.Label5.Name = "Label5";
            this.Label5.Size = new System.Drawing.Size(293, 53);
            this.Label5.TabIndex = 6;
            this.Label5.Text = "move,Solrock,Rapid Spin       trait,Gengar,Clear Body,2     illegal,Blissey,Heal Bell,Softboiled";
            // 
            // Label4
            // 
            this.Label4.Location = new System.Drawing.Point(148, 248);
            this.Label4.Name = "Label4";
            this.Label4.Size = new System.Drawing.Size(69, 13);
            this.Label4.TabIndex = 5;
            this.Label4.Text = "Examples:";
            // 
            // Label3
            // 
            this.Label3.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label3.Location = new System.Drawing.Point(152, 188);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(317, 49);
            this.Label3.TabIndex = 4;
            this.Label3.Text = "move,[pokémon],[attack]    trait,[pokémon],[trait],[slot] illegal,[pokémon],[move1],[move2],[...]";
            // 
            // Label2
            // 
            this.Label2.Location = new System.Drawing.Point(148, 172);
            this.Label2.Name = "Label2";
            this.Label2.Size = new System.Drawing.Size(69, 13);
            this.Label2.TabIndex = 3;
            this.Label2.Text = "Syntax:";
            // 
            // Label1
            // 
            this.Label1.Location = new System.Drawing.Point(4, 172);
            this.Label1.Name = "Label1";
            this.Label1.Size = new System.Drawing.Size(129, 89);
            this.Label1.TabIndex = 2;
            this.Label1.Text = "\u0000\u0000\u0000Enter Database changes here.  Changes will apply to users of this server only.  Teams made using these changes may not be legal on other server";
            // 
            // DatabaseChanges
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(472, 320);
            this.Controls.Add(this.Command2);
            this.Controls.Add(this.Command1);
            this.Controls.Add(this.txtDB);
            this.Controls.Add(this.Label5);
            this.Controls.Add(this.Label4);
            this.Controls.Add(this.Label3);
            this.Controls.Add(this.Label2);
            this.Controls.Add(this.Label1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Location = new System.Drawing.Point(3, 22);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DatabaseChanges";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Per-Server Database Changes";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Button Command2;
        private System.Windows.Forms.Button Command1;
        private System.Windows.Forms.TextBox txtDB;
        private System.Windows.Forms.Label Label5;
        private System.Windows.Forms.Label Label4;
        private System.Windows.Forms.Label Label3;
        private System.Windows.Forms.Label Label2;
        private System.Windows.Forms.Label Label1;
    }
}
