using System;
using System.Drawing;
using System.Windows.Forms;
using Netbattle.Common;
using Netbattle.Network;

namespace Netbattle.Forms {
    public partial class PrivateMessage : Form {
        private NbClient _myClient;
        public Player _playerTo;

        public PrivateMessage(Form parent, NbClient client, Player messagee) {
            this.MdiParent = parent;
            _myClient = client;
            _playerTo = messagee;
            InitializeComponent();
        }

        private void PrivateMessage_Load(object sender, EventArgs e) {
            this.Text = "Private Message: " + _playerTo.Name;
        }

        public void HandleIncoming(string message) {

            if (!message.StartsWith("/me ")) {
                AddMessage(_playerTo.Name + ":", Color.Blue, false, true);
                AddMessage(" " + message);
            } else if (message.StartsWith("/me ")) {
                AddMessage($"*** {_playerTo.Name} {message.Substring(4)}", Color.DarkOrchid);
            }
        }

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

        private void btnSend_Click(object sender, EventArgs e) {
            string message = txtInput.Text.Trim();
            txtInput.Clear();

            if (string.IsNullOrWhiteSpace(message))
                return;

            if (message.StartsWith("/me ")) {
                AddMessage($"*** {_myClient.You.Name} {message.Substring(4)}", Color.DarkOrchid);
            }
            else {
                AddMessage(_myClient.You.Name + ":", Color.Red, false, true);
                AddMessage(" " + message);
            }

            while (message.Length > 200) {
                _myClient.SendInstantMessage(message.Substring(0, 200), (byte)_playerTo.Id);
                message = message.Substring(200);
            }

            _myClient.SendInstantMessage(message, (byte)_playerTo.Id);
        }

        private void PrivateMessage_FormClosing(object sender, FormClosingEventArgs e) {

        }
    }
}
