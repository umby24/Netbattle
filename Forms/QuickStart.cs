using System;
using System.Windows.Forms;
using Netbattle.Forms;

namespace Netbattle {
    public partial class Form1 : Form {
        public Form1(Form parent) {
            this.MdiParent = parent;
            InitializeComponent();
        }

        private void btnJoinServer_Click(object sender, EventArgs e) {
            var browser = new ServerList(MdiParent);
            browser.FormClosing += BrowserOnFormClosed;
            browser.Closed += BrowserOnClosed;
            browser.Show();
            Hide();
        }

        private void BrowserOnClosed(object sender, EventArgs eventArgs) {
            Console.WriteLine("Browser closed2, showing!");
            Show();
        }

        private void BrowserOnFormClosed(object sender, FormClosingEventArgs formClosedEventArgs) {
            Console.WriteLine("Browser closed, showing!");
            Show();
        }

        private void btnTeamBuilder_Click(object sender, EventArgs e) {
            var tb = new TeamBuilder(MdiParent);
            tb.Show();
            //var meh = new ServerWindow(MdiParent, "127.0.0.1");
            //meh.Show();
        }

        private void Form1_Load(object sender, EventArgs e) {

        }

        private void btnAbout_Click(object sender, EventArgs e) {

        }

        private void btnExit_Click(object sender, EventArgs e) {
            Application.Exit();
        }
    }
}
