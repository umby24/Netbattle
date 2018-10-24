using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Netbattle.Common;
using Netbattle.Network;

namespace Netbattle.Forms {
    public partial class ServerWindow : Form {
        private string _serverIp;
        private NbClient _client;
        private Dictionary<int, PrivateMessage> _pmWindows = new Dictionary<int, PrivateMessage>();

        public ServerWindow(Form parent, string ip) {
            this.MdiParent = parent;
            InitializeComponent();
            _serverIp = ip;
        }

        #region Form Events
        private void ServerWindow_Load(object sender, EventArgs e) {
            _client = new NbClient(_serverIp);
            _client.ServerInfoReceived += ClientOnServerInfoReceived;
            _client.ChatMessageReceived += ClientOnChatMessageReceived;
            _client.WelcomeMessageReceived += ClientOnWelcomeMessageReceived;
            _client.PlayerlistUpdated += ClientOnPlayerlistUpdated;
            _client.PlayerJoined += ClientOnPlayerJoined;
            _client.PlayerLeft += ClientOnPlayerLeft;
            _client.PmReceived += ClientOnPmReceived;
            _client.Connect();
        }

        private void ClientOnPmReceived(Player player, string message) {
            if (InvokeRequired) {
                Invoke(new PrivateMessageEventArgs(ClientOnPmReceived), player, message);
                return;
            }

            PrivateMessage theWindow;

            if (!_pmWindows.ContainsKey(player.Id)) {
                theWindow = new PrivateMessage(this.MdiParent, _client, player);
                theWindow.FormClosing += TheWindowOnFormClosing;
                _pmWindows.Add(player.Id, theWindow);
                theWindow.Show();
            }
            else {
                theWindow = _pmWindows[player.Id];
            }

            theWindow.HandleIncoming(message);
        }

        private void ClientOnPlayerLeft(Player player) {
            if (InvokeRequired) {
                Invoke(new PlayerEventArgs(ClientOnPlayerLeft), player);
                return;
            }

            AddMessage($"{player.Name} left.");
        }

        private void ClientOnPlayerJoined(Player player) {
            if (InvokeRequired) {
                Invoke(new PlayerEventArgs(ClientOnPlayerJoined), player);
                return;
            }

            AddMessage($"{player.Name} joined.");
        }

        private void ClientOnPlayerlistUpdated(List<Player> players) {
            if (InvokeRequired) {
                Invoke(new PlayersEventArgs(ClientOnPlayerlistUpdated), players);
                return;
            }

            lstPlayers.Items.Clear();

            foreach (Player player in players) {
                string[] row = {player.Name};
                var lvi = new ListViewItem(row);
                lvi.ImageIndex = player.Picture - 1;
                lstPlayers.Items.Add(lvi);
            }
        }

        private void ClientOnWelcomeMessageReceived(string message) {
            AddMessage("Welcome Message: ", Color.Red, false, true, true);
            AddMessage(message);
        }

        private void ClientOnChatMessageReceived(string message) {
            if (message.Contains(":") && !message.StartsWith("*** ")) {
                Color nameColor = Color.Teal;
                var name = message.Substring(0, message.IndexOf(":"));

                if (name == _client.You.Name)
                    nameColor = Color.Red;
                
                if (_client.OnlinePlayers != null && _client.OnlinePlayers.Values.Any(a => a.Name == name && a.Name != _client.You.Name)) {
                    nameColor = Color.Blue;
                }
                
                AddMessage(name, nameColor, false, true);
                message = message.Substring(name.Length, message.Length - name.Length);
                AddMessage(message);
            } else if (message.StartsWith("*** ")) {
                AddMessage(message, Color.DarkOrchid);
            }
            else {

                AddMessage(message);
            }
        }


        private void btnSend_Click(object sender, EventArgs e) {
            _client.SendChat(txtInput.Text);
            txtInput.Clear();
        }
        #endregion

        #region Form Helpers

        public delegate void AddMessageArgs(
            string message, Color color = default(Color), bool newLine = true, bool bold = false, bool italic = false);

        public void AddMessage(string message, Color color = default(Color), bool newLine = true, bool bold = false, bool italic = false) {
            if (InvokeRequired) {
                Invoke(new AddMessageArgs(AddMessage), message, color, newLine, bold, italic);
                return;
            }

            if (newLine)
                message += "\r\n";

            txtChat.AppendText(message);
            if (newLine)
                txtChat.Select(txtChat.Text.Length - (message.Length - 1), message.Length - 1);
            else
                txtChat.Select(txtChat.Text.Length - message.Length, message.Length);

            txtChat.SelectionColor = color;

            FontStyle style = (bold ? FontStyle.Bold : 0) | (italic ? FontStyle.Italic : 0);
            
            txtChat.SelectionFont = new Font("MS Sans Serif", 8, style);
            txtChat.Select(txtChat.Text.Length, txtChat.Text.Length);
            txtChat.ScrollToCaret();
        }
        #endregion

        #region Event Handlers
        private void ClientOnServerInfoReceived(ServerInfo info) {
            if (InvokeRequired) {
                Invoke(new ServerInfoEventArgs(ClientOnServerInfoReceived), info);
                return;
            }

            AddMessage($"Server {info.ServerName} - NetBattle {info.ServerVersion}", Color.Blue, true, true);
            AddMessage($"Currently {info.OnlinePlayers} trainer(s) online, with a maximum of {info.MaxPlayers} trainers.", Color.Blue, true, true);
            AddMessage($"Flood count is set to {info.FloodTolerance}", Color.Red, true, true);

            if (_client.EncryptionEnabled)
                AddMessage("This server encrypts traffic.", Color.Green, true, true);

            this.Text = "Stadium: " + info.ServerName;
        }
        #endregion

        private void ServerWindow_FormClosing(object sender, FormClosingEventArgs e) {
            if (_client.Connected) {
                _client.Disconnect();
            }
        }

        private void privateMessageToolStripMenuItem_Click(object sender, EventArgs e) {
            if (lstPlayers.SelectedIndices.Count == 0)
                return;

            var selectedItem = lstPlayers.SelectedIndices[0];
            var player = _client.OnlinePlayers.ElementAt(selectedItem);

            if (player.Value.Name == _client.You.Name)
                return;

            PrivateMessage theWindow;

            if (!_pmWindows.ContainsKey(player.Value.Id)) {
                theWindow = new PrivateMessage(this.MdiParent, _client, player.Value);
                theWindow.FormClosing += TheWindowOnFormClosing;
                _pmWindows.Add(player.Value.Id, theWindow);
                theWindow.Show();
            }
        }

        private void TheWindowOnFormClosing(object sender, FormClosingEventArgs e) {
            _pmWindows.Remove( ((PrivateMessage)sender)._playerTo.Id);
        }
    }
}
