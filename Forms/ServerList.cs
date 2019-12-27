using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Netbattle.Common;
using Netbattle.Network;
using Sockets;
using Sockets.EventArgs;

namespace Netbattle.Forms {
    public partial class ServerList : Form {
        private const string RegIp = "registry.pmnb.net";
        private const int RegPort = 30002;
        private readonly ClientSocket _regSock;
        private readonly ByteBuffer _sendBuffer;
        private readonly ByteBuffer _receiveBuffer;
        private readonly bool _canReceive;
        private Dictionary<string, IRegPacket> _packets;
        public List<ServerListing> Servers;

        public ServerList(Form parent) {
            MdiParent = parent;
            InitializeComponent();
            _sendBuffer = new ByteBuffer();
            _receiveBuffer = new ByteBuffer();
            _sendBuffer.DataAdded += SendBufferOnDataAdded;
            Servers = new List<ServerListing>();
            PopulatePackets();

            _regSock = new ClientSocket(RegIp, RegPort);
            _regSock.DataReceived += ReceivedRegistryData;
            _regSock.Disconnected += RegistryDisconnected;
            _canReceive = true;

            try {
                _regSock.Connect();
            }
            catch (Exception ex) {
                MessageBox.Show("Failed to connect to netbattle Registry! It might be down!", "Registry down!",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void SendBufferOnDataAdded() {
            byte[] data = _sendBuffer.GetAllBytes();
            byte[] encrypted = XorModule.XorEncrypt(data);

            var length = (byte) (encrypted.Length - 1); // -- Because netbattle logic..
            var finalData = new[] {length}; // -- Prepend the finished array with the length.
            finalData = finalData.Concat(encrypted).ToArray();
            _regSock.Send(finalData);
        }

        public void RefreshList() {
            lstServers.Items.Clear();

            foreach (ServerListing serverListing in Servers) {
                string[] row = { serverListing.Name, serverListing.Owner, serverListing.OnlinePlayers + "/" + serverListing.MaxPlayers };
                lstServers.Items.Add(new ListViewItem(row));
            }
        }

        public void SendPacket(IRegPacket regPacket) {
            lock (_sendBuffer) {
                regPacket.Write(_sendBuffer);
            }
        }

        private void RegistryDisconnected(SocketDisconnectedArgs args) {
        }

        private void ReceivedRegistryData(DataReceivedArgs args) {
            if (!_canReceive)
                return;

            lock (_receiveBuffer) {
                _receiveBuffer.AddBytes(args.Data);
            }

            HandlePackets();
        }

        private void btnCancel_Click(object sender, EventArgs e) {
            Close();
        }

        private void PopulatePackets() {
            _packets = new Dictionary<string, IRegPacket> {
                {"SERV", new ServerRegPacket() },
                {"TBAN", new TempbanRegPacket() },
                {"MULTI", new DupeRegPacket() },
                {"PING", new PingRegPacket() },
                {"DISC", new MoveToTopRegPacket() }
            };
        }

        public void HandlePackets() {
            lock (_receiveBuffer) {
                if (_receiveBuffer.Length == 0)
                    return;

                var length = (byte) (_receiveBuffer.PeekByte() + 1);

                if (_receiveBuffer.Length - 1 < length)
                    return;

                _receiveBuffer.ReadByte(); // -- Dispose of the length code..
                // -- Decrypt the data (Registry is always encrypted)
                byte[] data = XorModule.XorDecrypt(_receiveBuffer.ReadByteArray(length));

                var tempBuffer = new ByteBuffer(); // -- Create a new bytebuffer with the decrypted data
                tempBuffer.AddBytes(data);

                string cmd = tempBuffer.ReadString(4); // -- Get the command to find the regPacket parser..
                tempBuffer.ReadByte(); // -- Trim off the colon.

                // -- try to find the regPacket.
                IRegPacket regPacket;
                if (!_packets.TryGetValue(cmd, out regPacket)) {
                    MessageBox.Show("Invalid Packet recieved!!");
                    return;
                }

                regPacket.Read(tempBuffer);
                regPacket.Handle(this);
            }
        }

        private void lstServers_SelectedIndexChanged(object sender, EventArgs e) {
            if (lstServers.SelectedIndices.Count == 0)
                return;

            if (lstServers.SelectedIndices[0] > Servers.Count - 1)
                return;

            txtDescription.Text = Servers[lstServers.SelectedIndices[0]].Description;
        }

        private void btnConnect_Click(object sender, EventArgs e) {
            if (lstServers.SelectedIndices.Count == 0)
                return;

            if (lstServers.SelectedIndices[0] > Servers.Count - 1)
                return;

            string ip = UnpackIp(Servers[lstServers.SelectedIndices[0]].Ip);

            if (ip == "66.69.46.171")
                ip = "127.0.0.1";

            var srv = new ServerWindow(MdiParent, ip);
            srv.FormClosing += Srv_FormClosed;
            srv.Show();
            Hide();
        }

        private void Srv_FormClosed(object sender, FormClosingEventArgs e) {
            if (this.InvokeRequired) {
                Invoke(new FormClosingEventHandler(Srv_FormClosed), sender, e);
                return;
            }

            Invoke(new NoEventArgs(derpish));
            
        }

        private void derpish() {
            this.Close();
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

        }
    }
}
