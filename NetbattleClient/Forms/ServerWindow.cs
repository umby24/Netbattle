using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Netbattle.Common;
using Netbattle.Core.Common;
using Netbattle.Core.Network;
using Netbattle.Network;
using BanResult = Netbattle.Core.Common.BanResult;

namespace Netbattle.Forms {
    public partial class ServerWindow : Form
    {
        private string _serverIp;
        private NbClient _client;
        private Dictionary<int, PrivateMessage> _pmWindows = new Dictionary<int, PrivateMessage>();
        private Dictionary<int, BattleWindow> _watchWindows = new Dictionary<int, BattleWindow>();
        private BattleWindow? _battleWindow;
        private UserInfo? _userInfoWindow;
        
        public ServerWindow(Form parent, string ip)
        {
            this.MdiParent = parent;
            InitializeComponent();
            MdiMenu.MergeIntoParent(this, menuStrip1);
            _serverIp = ip;
            var font = new Font("MS Sans Serif", 8.25f);
            lstPlayers.Font = font;
            txtChat.Font = font;
            txtInput.Font = font;
        }

        #region Form Events
        private void ServerWindow_Load(object sender, EventArgs e)
        {
            _client = new NbClient(_serverIp);
            _client.ServerInfoReceived += ClientOnServerInfoReceived;
            _client.ChatMessageReceived += ClientOnChatMessageReceived;
            _client.WelcomeMessageReceived += ClientOnWelcomeMessageReceived;
            _client.PlayerlistUpdated += ClientOnPlayerlistUpdated;
            _client.PlayerJoined += ClientOnPlayerJoined;
            _client.PlayerLeft += ClientOnPlayerLeft;
            _client.PlayerAway += ClientOnPlayerAway;
            _client.PlayerBack += ClientOnPlayerBack;
            _client.PlayerInfoUpdated += ClientOnPlayerInfoUpdated;
            _client.PlayerKickedByServer += ClientOnPlayerKickedByServer;
            _client.PlayerBanned += _client_PlayerBanned;
            _client.PlayerKicked += ClientOnPlayerKicked;
            _client.PlayerTempBanned += _client_PlayerTempBanned;
            _client.PlayerPingUpdated += ClientOnPlayerPingUpdated;
            _client.PmReceived += ClientOnPmReceived;
            _client.ServerKickedYou += ClientOnServerKickedYou;
            _client.DuplicateNameKick += ClientOnDuplicateNameKick;
            _client.InvalidUserPassword += ClientOnInvalidUserPassword;
            _client.IpBanned += ClientOnIpBanned;
            _client.ChallengeReceived += ClientOnChallengeReceived;
            _client.ChallengeAccepted += ClientOnChallengeAccepted;
            _client.ChallengeRefused += ClientOnChallengeRefused;
            _client.ChallengeEnded += ClientOnChallengeEnded;
            _client.PasswordRequested += ClientOnPasswordRequested;
            _client.RegistryMessageReceived += ClientOnRegistryMessageReceived;
            _client.TempBanned += ClientOnTempBanned;
            _client.ServerLocked += ClientOnServerLocked;
            _client.ServerFull += ClientOnServerFull;
            _client.ServerClosing += ClientOnServerClosing;
            _client.BattleStarted += _client_BattleStarted;
            _client.BattleStart += ClientOnBattleStart;
            _client.Connect();
        }

        private void ClientOnBattleStart(string message) {
            openTeamBuilderToolStripMenuItem.Enabled = false;
            if (_battleWindow != null) {
                _battleWindow.Close();
                _battleWindow.Dispose();
            }

            _battleWindow = new BattleWindow();
            _battleWindow.Show();
        }

        private void ClientOnServerClosing() {
            AddMessage("SERVER IS SHUTTING DOWN", Color.Red, true, true);
            _client.Disconnect();
        }

        private void ClientOnServerFull() {
            MessageBox.Show("Server is full - try again later.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            _client.Disconnect();
        }

        private void ClientOnPlayerKicked(Player player1, Player player2) {
            var authString = "";
            if (player1.Authority == 1)
            {
                authString = "Moderator";
            } else
            {
                authString = "Admin";
            }

            AddMessage($"{authString} {player1.Name} has kicked {player2.Name}.", Color.Red, true, true);
        }

        private void ClientOnPlayerPingUpdated(Player player) {
            ChallengeWindow.Current?.Text = ChallengeWindow.Current?.Text + " - Ping Speed: " + player.Ping;
        }

        private void ClientOnChallengeEnded() {
            ChallengeWindow.Current?.CloseAndCancel();
            _client.ChallengeNumber = 0;
            _client.ChallengePending = false;
            _client.ChallengeWindowOpen = false;
            _client.ChallengeIsOutgoing = false;
        }

        private void ClientOnChallengeRefused(Player player) {
            // -- Noop..
        }

        private void ClientOnChallengeAccepted(Player player) {
            lblStatus.Text = "Challenge Accepted!";
        }

        private void _client_BattleStarted(Player player1, Player player2)
        {
            AddMessage($"{player1.Name} and {player2.Name} are battling.");
        }

        private void _client_PlayerTempBanned(Player player1, Player player2)
        {
            var authString = "";
            if (player1.Authority == 1)
            {
                authString = "Moderator";
            } else
            {
                authString = "Admin";
            }

            AddMessage($"{authString} {player1.Name} has TempBanned {player2.Name}.", Color.Red, true, true);
        }

        private void _client_PlayerBanned(Player player1, Player player2)
        {
            AddMessage($"Admin {player1.Name} has banned " + player2.Name, Color.Red, true, true);
        }

        private void ClientOnPlayerKickedByServer(Player player)
        {
            AddMessage($"SERVER HAS KICKED TRAINER: {player.Name}", Color.Red, true, true);
        }

        private void ClientOnPlayerInfoUpdated(Player player)
        {
            if (InvokeRequired)
            {
                Invoke(new PlayerEventArgs(ClientOnPlayerInfoUpdated), player);
                return;
            }

            if (player.Id == _client.You.Id && player.Authority > 1) {
                controlWindowToolStripMenuItem.Visible = true;
                kickToolStripMenuItem.Visible = true;
            } else if (player.Id == _client.You.Id) {
                controlWindowToolStripMenuItem.Visible = false;
                kickToolStripMenuItem.Visible = false;
            }
            RefreshPlayerList();
        }

        private void ClientOnServerLocked() {
            MessageBox.Show("This server is not accepting new Trainers at this time. Please try again later.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            _client.Disconnect();
        }

        private void ClientOnTempBanned(string message) {
            MessageBox.Show("You have been TempBanned! - " + message, "Banned", MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            _client.Disconnect();
        }

        private void ClientOnRegistryMessageReceived(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new MessageEventArgs(ClientOnRegistryMessageReceived), message);
                return;
            }

            AddMessage("NETWORK-WIDE MESSAGE: ", ColorTranslator.FromOle(0x80ff), false, true);
            AddMessage(message);
        }

        private void ClientOnPasswordRequested()
        {
            var passwordForm = new ServerPassword();
            if (passwordForm.ShowDialog() == DialogResult.OK)
            {
                _client.SendServerPassword(passwordForm.Password);
            }
            else
            {
                _client.Disconnect();
                Close();
            }
        }

        private void ClientOnChallengeReceived(Player challenger, ChallengeSettings settings)
        {
            if (InvokeRequired)
            {
                Invoke(new ChallengeEventArgs(ClientOnChallengeReceived), challenger, settings);
                return;
            }

            _client.ChallengeNumber = challenger.Id;
            _client.ChallengeWindowOpen = true;
            _client.ChallengeIsOutgoing = true;
            _client.ChallengePending = true;
            
            var chWind = new ChallengeWindow(_client, challenger, this, settings);
            chWind.Show();
        }

        private void ClientOnPlayerBack(Player player)
        {
            if (InvokeRequired)
            {
                Invoke(new PlayerEventArgs(ClientOnPlayerBack), player);
                return;
            }

            if (player.BattlingWith == 1025)
            {
                AddMessage($"{player.Name} has returned.");
            }
            else
            {
                AddMessage($"{player.Name} is done battling.");
                // -- If Watching, drop the watching form.
                if (player.BattlingWith == _client.You.Id && _client.OnlinePlayers[_client.You.Id].Id == player.Id)
                {
                    player.BattlingWith = 0;
                    // -- Call "BOVER" to battle window.
                }
            }

            player.BattlingWith = 0;
            RefreshPlayerList();
        }

        private void ClientOnPlayerAway(Player player)
        {
            if (InvokeRequired)
            {
                Invoke(new PlayerEventArgs(ClientOnPlayerAway), player);
                return;
            }

            AddMessage($"{player.Name} is away.");
            player.BattlingWith = 1025;
            RefreshPlayerList();
        }

        private void ClientOnIpBanned()
        {
            AddMessage("You have been banned from this server.", Color.Red, true, true, true);
        }

        private void ClientOnInvalidUserPassword()
        {
            AddMessage("Your user password is invalid. Try another.", Color.Red, true, true, true);
        }

        private void ClientOnDuplicateNameKick()
        {
            AddMessage("Name Already In Use", Color.Red, true, true, true);
        }

        private void ClientOnServerKickedYou()
        {
            MessageBox.Show("You have been kicked by server " + _client.Info.ServerName, "Server Kicked You", MessageBoxButtons.OK, MessageBoxIcon.Error);
            AddMessage("Disconnected from server", Color.Red, true, true, true);
        }

        private void ClientOnPmReceived(Player player, string message)
        {
            if (InvokeRequired)
            {
                Invoke(new PrivateMessageEventArgs(ClientOnPmReceived), player, message);
                return;
            }

            PrivateMessage theWindow;

            if (!_pmWindows.ContainsKey(player.Id))
            {
                theWindow = new PrivateMessage(this.MdiParent, _client, player);
                theWindow.FormClosing += TheWindowOnFormClosing;
                _pmWindows.Add(player.Id, theWindow);
                theWindow.Show();
            }
            else
            {
                theWindow = _pmWindows[player.Id];
            }

            theWindow.HandleIncoming(message);
        }

        private void ClientOnPlayerLeft(Player player)
        {
            if (InvokeRequired)
            {
                Invoke(new PlayerEventArgs(ClientOnPlayerLeft), player);
                return;
            }

            lblStatus.Text = "Trainer signed off";
            if (player.BattlingWith > 0 && player.BattlingWith < 1025) {
                AddMessage($"{player.Name} has disconnected from battle.");
            }
            
            AddMessage($"{player.Name} signed off.");
            
            if (_pmWindows.ContainsKey(player.Id)) {
                _pmWindows[player.Id].Close();
            }

            if (_client.ChallengeNumber == player.Id) {
                _client.ChallengeNumber = 0;
                _client.ChallengePending = false;
                _client.ChallengeWindowOpen = false;
                ChallengeWindow.Current?.CloseAndCancel();
            }

            if (_client.You.BattlingWith == player.Id) {
                // -- BattleDisconnectIcon
            }
            // -- If you have a watch window for this person, OppDiscon()
            
        }

        private void ClientOnPlayerJoined(Player player)
        {
            if (InvokeRequired)
            {
                Invoke(new PlayerEventArgs(ClientOnPlayerJoined), player);
                return;
            }

            AddMessage($"{player.Name} joined.");
        }

        private void ClientOnPlayerlistUpdated(List<Player> players)
        {
            if (InvokeRequired)
            {
                Invoke(new PlayersEventArgs(ClientOnPlayerlistUpdated), players);
                return;
            }

            lblStatus.Text = "Received Trainer List";
            RefreshPlayerList();
        }

        private void ClientOnWelcomeMessageReceived(string message)
        {
            AddMessage("Welcome Message: ", Color.Red, false, true, true);
            AddMessage(message);
        }

        private void ClientOnChatMessageReceived(string message)
        {
            if (message.Contains(":") && !message.StartsWith("*** "))
            {
                Color nameColor = ColorTranslator.FromOle(0x99AA00);
                string name = message.Substring(0, message.IndexOf(":"));

                if (name == _client.You.Name)
                    nameColor = Color.Red;

                if (_client.OnlinePlayers.Values.Any(a => a.Name == name && a.Name != _client.You.Name))
                {
                    nameColor = Color.Blue;
                }

                AddMessage(name, nameColor, false, true);
                message = message.Substring(name.Length, message.Length - name.Length);
                AddMessage(message);
            }
            else if (message.StartsWith("*** "))
            {
                AddMessage(message, ColorTranslator.FromOle(0xC000C0));
            }
            else
            {

                AddMessage(message);
            }
        }


        private void btnSend_Click(object sender, EventArgs e)
        {
            _client.SendChat(txtInput.Text);
            txtInput.Clear();
        }
        #endregion

        #region Form Helpers

        public void RefreshPlayerList()
        {
            lstPlayers.Items.Clear();

            foreach (Player player in _client.OnlinePlayers.Values)
            {
                var displayName = player.Away ? "[" + player.Name + "]" : player.Name;

                if (player.Authority == 2)
                    displayName += " *";

                if (player.Authority == 3)
                    displayName += " **";

                string[] row = { displayName };

                // -- Pictures are 1-based, like the VB ListImages they index.
                var lvi = new ListViewItem(row)
                {
                    ImageIndex = player.Picture - 1,
                    ForeColor = player.Away ? Color.Gray : Color.Black
                };

                lstPlayers.Items.Add(lvi);
            }
        }

        public delegate void AddMessageArgs(
            string message, Color color = default(Color), bool newLine = true, bool bold = false, bool italic = false);

        public void AddMessage(string message, Color color = default(Color), bool newLine = true, bool bold = false, bool italic = false)
        {
            if (InvokeRequired)
            {
                Invoke(new AddMessageArgs(AddMessage), message, color, newLine, bold, italic);
                return;
            }

            if (newLine)
                message += "\r\n";
            
            // -- Apply timestamps but only if its the beginning of a new line..
            if (Configuration.CurrentSettings.UseTimestamp && txtChat.Text.Length > 0 && txtChat.Text.Substring(txtChat.Text.Length - 1) == "\n") {
                var tsString = DateTime.Now.ToString(Configuration.CurrentSettings.TimestampFormat);
                txtChat.AppendText(tsString + " ");
            }
            
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
        private void ClientOnServerInfoReceived(ServerInfo info)
        {
            if (InvokeRequired)
            {
                Invoke(new ServerInfoEventArgs(ClientOnServerInfoReceived), info);
                return;
            }

            AddMessage($"Server {info.ServerName} - NetBattle {info.ServerVersion}", Color.Blue, true, true);
            AddMessage($"Currently {info.OnlinePlayers} trainer(s) online, with a maximum of {info.MaxPlayers} trainers.", Color.Blue, true, true);
            AddMessage($"Flood count is set to {info.FloodTolerance}", Color.Red, true, true);
            AddMessage($"Your team's power is {_client.You.TeamPower}.", Color.Black, true, true);

            if (_client.EncryptionEnabled)
                AddMessage("This server encrypts traffic.", Color.Green, true, true);

            if (_client.You.Authority == 0) {
                controlWindowToolStripMenuItem.Visible = false;
                kickToolStripMenuItem.Visible = false;
            }
            this.Text = "Stadium: " + info.ServerName;
            
        }
        #endregion

        private void ServerWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_client.Connected)
            {
                _client.Disconnect();
            }
        }

        private void privateMessageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lstPlayers.SelectedIndices.Count == 0)
                return;

            var selectedItem = lstPlayers.SelectedIndices[0];
            var player = _client.OnlinePlayers.ElementAt(selectedItem);

            if (player.Value.Name == _client.You.Name)
                return;

            PrivateMessage theWindow;

            if (!_pmWindows.ContainsKey(player.Value.Id))
            {
                theWindow = new PrivateMessage(this.MdiParent, _client, player.Value);
                theWindow.FormClosing += TheWindowOnFormClosing;
                _pmWindows.Add(player.Value.Id, theWindow);
                theWindow.Show();
            }
        }

        private void TheWindowOnFormClosing(object sender, FormClosingEventArgs e)
        {
            _pmWindows.Remove(((PrivateMessage)sender)._playerTo.Id);
        }

        public void ToggleAway()
        {
            awayToolStripMenuItem.Checked = !awayToolStripMenuItem.Checked;
            if (awayToolStripMenuItem.Checked)
            {
                _client.You.Away = true;
                _client.SendPacket(new PlayerAway());
            }
            else
            {
                _client.You.Away = false;
                _client.SendPacket(new PlayerBack());
            }
        }

        private void awayToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ToggleAway();
        }

        private void openTeamBuilderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!_client.You.Away)
                ToggleAway();

            var tb = new TeamBuilder(this);
            tb.FormClosed += (s, args) =>
            {
                if (_client.You.Away)
                    ToggleAway();
            };
            tb.Show();
        }

        private void controlWindowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lstPlayers.SelectedIndices.Count == 0)
                return;

            var selectedItem = lstPlayers.SelectedIndices[0];
            var player = _client.OnlinePlayers.ElementAt(selectedItem);

            if (_userInfoWindow != null) {
                _userInfoWindow.Close();
                _userInfoWindow.Dispose();
                _userInfoWindow = null;
            }
            
            _userInfoWindow = new UserInfo(this.MdiParent, _client, player.Value);
            _userInfoWindow.Show();
        }

        private void kickToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void challengeInfoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lstPlayers.SelectedIndices.Count == 0)
                return;

            var selectedItem = lstPlayers.SelectedIndices[0];
            var player = _client.OnlinePlayers.ElementAt(selectedItem);

            if (player.Value.Name == _client.You.Name)
                return;

            ChallengeWindow.Current?.CloseAndCancel(); // -- Only one challenge at a time, as in VB.
            _client.ChallengeWindowOpen = true;
            _client.ChallengeNumber = player.Value.Id;
            var chln = new ChallengeWindow(_client, player.Value, this);
            chln.Show();
        }

        private void watchBattleToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
