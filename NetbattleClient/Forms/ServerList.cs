using Netbattle;
using Netbattle.Core.Common;
using Netbattle.Core.Network;
using Netbattle.Core.Database;
using Netbattle.Forms;

namespace NetbattleClient.Forms {
    public partial class ServerList : Form {
        private RegistryClient _regClient;

        /// <summary>
        /// The server window opened from this list, if the user connected to one.
        /// </summary>
        public ServerWindow? OpenedServer { get; private set; }

        public ServerList(Form parent) {
            MdiParent = parent;
            InitializeComponent();
            lblStatus.Text = "Connecting to Registry at " + Configuration.CurrentSettings.RegistryAddress;
            try {
                _regClient = new RegistryClient();
                _regClient.Connected += RegClientOnConnected;
                _regClient.ServerListChanged += RefreshList;
                _regClient.Disconnected += RegClientOnDisconnected;
                _regClient.RegistryBanned += RegClientOnRegistryBanned;
            }
            catch (Exception ex) {
                MessageBox.Show("Failed to connect to server registry! Trying advanced mode.");
                btnDirect.PerformClick();
            }
        }

        private void RegClientOnConnected() {
            lblStatus.Text = "Connected!";
        }

        private void RegClientOnRegistryBanned() {
            MessageBox.Show("You have been banned from accessing the Netbattle Registry.", "Banned",
                MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void RegClientOnDisconnected() {
            lblStatus.Text = "Disconnected from Registry.";
        }


        private void btnCancel_Click(object sender, EventArgs e) {
            Close();
        }
        public void RefreshList() {
            if (InvokeRequired) {
                Invoke(new MethodInvoker(RefreshList));
                return;
            }

            lblStatus.Text = "Connected!";
            
            lstServers.Items.Clear();

            foreach (ServerListing serverListing in _regClient.Servers) {
                string[] row = { serverListing.Name, serverListing.Owner, serverListing.OnlinePlayers + "/" + serverListing.MaxPlayers };
                lstServers.Items.Add(new ListViewItem(row));
            }
        }
        private void lstServers_SelectedIndexChanged(object sender, EventArgs e) {
            if (lstServers.SelectedIndices.Count == 0)
                return;

            if (lstServers.SelectedIndices[0] > _regClient.Servers.Count - 1)
                return;

            txtDescription.Text = _regClient.Servers[lstServers.SelectedIndices[0]].Description;
        }

        private void btnConnect_Click(object sender, EventArgs e) {
            if (lstServers.SelectedIndices.Count == 0)
                return;

            if (lstServers.SelectedIndices[0] > _regClient.Servers.Count - 1)
                return;

            string ip = UnpackIp(_regClient.Servers[lstServers.SelectedIndices[0]].Ip);

            // -- Close rather than Hide: hiding an MDI child destroys its handle, and a later Close() then disposes it
            // -- without raising FormClosing/FormClosed. Whoever opened us picks the server window up from OpenedServer.
            OpenedServer = new ServerWindow(MdiParent, ip);
            OpenedServer.Show();
            Close();
        }

        private string UnpackIp(string ip) {
            int first = Convert.ToInt32(ip.Substring(0, 2), 16);
            int second = Convert.ToInt32(ip.Substring(2, 2), 16);
            int third = Convert.ToInt32(ip.Substring(4, 2), 16);
            int forth = Convert.ToInt32(ip.Substring(6, 2), 16);

            return $"{first}.{second}.{third}.{forth}";
        }

        private void ServerList_Load(object sender, EventArgs e) {
            lstServers.FullRowSelect = true;
        }

        private void ServerList_FormClosed(object sender, FormClosedEventArgs e) {
            if (_regClient != null) {
                _regClient.ServerListChanged -= RefreshList;
                _regClient.Disconnect();
            }
        }

        private void lstServers_MouseDoubleClick(object sender, MouseEventArgs e) {
            btnConnect.PerformClick();
        }

        private void btnDirect_Click(object sender, EventArgs e) {
            var advWindow = new AdvancedConnect(MdiParent);
            advWindow.Show();
            Close();
        }
    }
}
