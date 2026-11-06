using System.Windows.Forms;
using Netbattle.Core.Common;
using Netbattle.Core.Network;

namespace Netbattle.Forms
{
    // -- Layout converted from PlayerInfoAdv.frm. None of its code has been ported yet.
    public partial class UserInfo : Form
    {
        private readonly Player _targetPlayer;
        private readonly NbClient _client;
        private bool _playerOnline;
        private string _lookingUp = ""; // -- The name we're waiting on a LOOK: answer for, as in VB's LookingUp.
        public UserInfo(Form mdiParent, NbClient client, Player targetPlayer)
        {
            InitializeComponent();
            this.MdiParent = mdiParent;
            _targetPlayer = targetPlayer;
            _client = client;

            // -- Subscribe once here rather than on every button click, so handlers don't stack up.
            _client.LookupResultReceived += _client_LookupResultReceived;
            _client.BanListUpdated += _client_BanListUpdated;
            _client.AliasDataRecieved += _client_HandleAliasData;

            FormClosed += (_, _) => {
                _client.LookupResultReceived -= _client_LookupResultReceived;
                _client.BanListUpdated -= _client_BanListUpdated;
                _client.AliasDataRecieved -= _client_HandleAliasData;
            };
            Text1.Text = targetPlayer.Name;
            Text2.Text = targetPlayer.Name;
        }

        private void _client_HandleAliasData(List<string> aliases)
        {
            if (InvokeRequired)
            {
                Invoke(new AliasDataEventArgs(_client_HandleAliasData), aliases);
                return;
            }

            NickList.Items.Clear();
            foreach (var alias in aliases)
            {
                NickList.Items.Add(alias);
            }
            cmdALook.Enabled = true;
            Command_0.Enabled = true;
            Command_1.Enabled = true;
            Command_2.Enabled = true;
            cmdLookup.Enabled = true;
            StatusBar_Panel1.Text = "Alias data received.";
        }

        private void UserInfo_Load(object sender, EventArgs e)
        {
            Command_1.Visible = (_client.You.Authority == 3);
            Command_2.Visible = (_client.You.Authority == 3);
            BanList.Items.Add("[Please click refresh]");
            cmdLookup.PerformClick();
        }

        private void cmdLookup_Click(object sender, EventArgs e)
        {
            cmdLookup.Enabled = false;
            Command_0.Enabled = false;
            Command_1.Enabled = false;
            Command_2.Enabled = false;
            txtInfo_0.Text = "[No User Selected]";
            txtInfo_1.Text = "";
            txtInfo_2.Text = "";
            txtInfo_3.Text = "";
            txtInfo_4.Text = "";
            StatusBar_Panel1.Text = "Querying, please wait...";
            _lookingUp = Text1.Text;
            _playerOnline = false;

            _client.QueryPlayerInfo(Text1.Text);

            foreach (var p in _client.OnlinePlayers.Values)
            {
                if (p.Name == Text1.Text)
                {
                    _playerOnline = true;
                    Command_1.Enabled = (p.Authority < _client.You.Authority);
                    break;
                }
            }
        }

        private void cmdALook_Click(object sender, EventArgs e)
        {
            cmdLookup.Enabled = false;
            cmdALook.Enabled = false;
            Command_0.Enabled = false;
            Command_1.Enabled = false;
            Command_2.Enabled = false;
            NickList.Items.Clear();
            _client.QueryPlayerAlias(Text2.Text);
            StatusBar_Panel1.Text = "Querying, please wait...";
        }

        private void Command_0_Click(object sender, EventArgs e)
        {
            var tbDuration = new TempBanSet();
            tbDuration.ShowDialog();
            if (tbDuration.DialogResult == DialogResult.OK)
            {
                _client.TempBan(Text1.Text, tbDuration.Value);
                StatusBar_Panel1.Text = "TempBan command sent.";
            }
        }

        private void Command_2_Click(object sender, EventArgs e)
        {
            _client.SidBan(Text1.Text);
            StatusBar_Panel1.Text = "SID Ban command sent.";
        }

        private void cmdUnban_Click(object sender, EventArgs e)
        {
            cmdRefresh.Enabled = false;
            cmdUnban.Enabled = false;
            StatusBar_Panel1.Text = "Unban command sent.";
            _client.Unban((string)BanList.SelectedItems[0].Tag, BanList.SelectedItems[0].Text);
        }

        private void Command_1_Click(object sender, EventArgs e)
        {
            _client.IPBan(Text1.Text);
            StatusBar_Panel1.Text = "IP Ban command sent.";
        }

        private void cmdRefresh_Click(object sender, EventArgs e)
        {
            cmdRefresh.Enabled = false;
            cmdUnban.Enabled = false;
            StatusBar_Panel1.Text = "Requesting ban list...";
            _client.RequestBanList();
        }

        private void _client_LookupResultReceived(UserLookupResult result)
        {
            if (InvokeRequired)
            {
                Invoke(new UserLookupEventArgs(_client_LookupResultReceived), result);
                return;
            }

            if (_lookingUp == "")
                return; // -- Not waiting on a lookup.

            if (!result.Found)
            {
                StatusBar_Panel1.Text = "Server reports no such user.";
                cmdLookup.Enabled = true;
                return;
            }

            if (!string.Equals(_lookingUp, result.Name, StringComparison.OrdinalIgnoreCase))
                return; // -- An answer to some other query.

            _lookingUp = result.Name;
            txtInfo_0.Text = result.Name;
            txtInfo_2.Text = result.StatusText;

            // -- QUIRK: VB compares your Authority (0-3) against the lookup's status (0-4), as is.
            bool canAct = _client.You.Authority >= result.Status;
            Command_0.Enabled = canAct && result.Status > 1; // -- TempBan
            Command_2.Enabled = canAct && result.Status > 0; // -- SID Ban

            txtInfo_3.Text = result.StationId;
            txtInfo_1.Text = _playerOnline ? "Currently Online" : "Last Online on " + result.LastLogin;

            if (_playerOnline)
            {
                txtInfo_4.Text = result.IP;
                Command_1.Enabled = canAct; // -- IP Ban
            }
            else
            {
                txtInfo_4.Text = "[Player not connected]";
                Command_1.Enabled = false;
            }

            cmdLookup.Enabled = true;
            cmdLookup.Focus();
            StatusBar_Panel1.Text = "Query successful.";
        }

        private void _client_BanListUpdated()
        {
            if (InvokeRequired)
            {
                Invoke(new EmptyEventArgs(_client_BanListUpdated));
                return;
            }

            BanList.Items.Clear();

            foreach(var item in _client.BanList)
            {
                var listItem = new ListViewItem(item.Name);
                listItem.Tag = item.Type;
                BanList.Items.Add(listItem);
            }
            cmdRefresh.Enabled = true;
            cmdUnban.Enabled = true;
            StatusBar_Panel1.Text = "Ban list received.";
        }
    }
}
