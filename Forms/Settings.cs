using System;
using System.Windows.Forms;
using Netbattle.Database;

namespace Netbattle.Forms
{
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();
        }

        private void txtRegAddr_TextChanged(object sender, EventArgs e)
        {
            UserSettings.CurrentSettings.RegistryAddress = txtRegAddr.Text;
        }

        private void Settings_Load(object sender, EventArgs e)
        {
            txtRegAddr.Text = UserSettings.CurrentSettings.RegistryAddress;
        }
    }
}