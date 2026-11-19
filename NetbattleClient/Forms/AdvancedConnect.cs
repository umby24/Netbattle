using Netbattle;
using Netbattle.Core.Database;
using Netbattle.Forms;

namespace NetbattleClient.Forms
{
    // -- Layout converted from NetSet.frm. None of its code has been ported yet.
    public partial class AdvancedConnect : Form
    {
        public ServerWindow? OpenedServer { get; private set; }
        
        public AdvancedConnect(Form parent) {
            this.MdiParent = parent;
            InitializeComponent();
        }

        private void Command1_Click(object sender, EventArgs e) {
            if (Mode_0.Checked) { // -- Spawn server list.
                var browser = new ServerList(MdiParent);
                browser.FormClosed += BrowserOnFormClosed;
                browser.Show();
                Hide();
                return;
            }

            if (!Configuration.CurrentSettings.AdvConnects.Contains(Address.Text)) {
                Configuration.CurrentSettings.AdvConnects.Add(Address.Text);
                Configuration.Save();
            }
            
            OpenedServer = new ServerWindow(MdiParent, Address.Text);
            OpenedServer.Show();
            Close();
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

        private void AdvancedConnect_Load(object sender, EventArgs e) {
            foreach (string addy in Configuration.CurrentSettings.AdvConnects) {
                Address.Items.Add(addy);
            }
        }

        private void Command2_Click(object sender, EventArgs e) {
            Close();
        }

        private void Command3_Click(object sender, EventArgs e) {
            Configuration.CurrentSettings.AdvConnects.Clear();
            Configuration.Save();
        }
    }
}
