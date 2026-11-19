using System;
using System.Windows.Forms;

namespace Netbattle.Forms {
    public partial class ServerPassword : Form {
        public string Password => txtPassword.Text;
        
        public ServerPassword() {
            InitializeComponent();
        }

        public void SetupSavedPw() {
            label1.Text = "Enter the user password to be saved.";
        }
        
        private void btnOk_Click(object sender, EventArgs e) {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void button2_Click(object sender, EventArgs e) {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}