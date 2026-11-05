using Netbattle.Forms;

namespace NetbattleClient.Forms {
    public partial class Form1 : Form {
        public Form1(Form parent) {
            this.MdiParent = parent;
            InitializeComponent();
        }

        private void btnJoinServer_Click(object sender, EventArgs e) {
            var browser = new ServerList(MdiParent);
            browser.FormClosed += BrowserOnFormClosed;
            browser.Show();
            Hide();
        }

        private void BrowserOnFormClosed(object sender, FormClosedEventArgs e) {
            ServerWindow? server = ((ServerList)sender).OpenedServer;

            if (server == null) {
                Show();
                return;
            }

            // -- Back to the quick start once the server window closes (it disconnects itself while closing).
            server.FormClosed += (_, _) => Show();
        }

        private void btnTeamBuilder_Click(object sender, EventArgs e) {
            var tb = new TeamBuilder(MdiParent);
            tb.Show();
        }

        private void Form1_Load(object sender, EventArgs e) {

        }

        private void btnAbout_Click(object sender, EventArgs e)
        {
            var about = new About();
            about.ShowDialog();
        }

        private void btnExit_Click(object sender, EventArgs e) {
            Application.Exit();
        }

        private void settingsToolStripMenuItem_Click(object sender, EventArgs e) {
            var settings = new Settings();
            settings.ShowDialog();
        }
    }
}
