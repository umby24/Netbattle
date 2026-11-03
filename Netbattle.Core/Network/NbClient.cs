using System.Reflection;
using Netbattle.Common;
using Netbattle.Core.Common;
using Netbattle.Core.Database;
using Netbattle.Network;
using Sockets;
using Sockets.EventArgs;

namespace Netbattle.Core.Network {
    public class NbClient {
        public Player You { get; set; }
        public string ServerIp { get; set; }
        public bool EncryptionEnabled { get; set; }
        public bool Connected { get; set; }
        
        public bool Battling { get; set; }
        
        public Dictionary<int, Player> OnlinePlayers { get; set; }
        #region Private Variables
        private ClientSocket _serverSocket;
        private ByteBuffer _inputBuffer;
        private ByteBuffer _outBuffer;
        private ByteBuffer _plistBuffer;
        private Dictionary<string, IPacket> _packets;
        public byte[] LoginCode { get; set; }
        #endregion

        #region Server Info 
        // -- TODO: Make this a class or something.
        public ServerInfo Info;
        #endregion

        public NbClient(string ip) {
            You = new Player {
                Name = "BOT"
            };
            ServerIp = ip;
            // -- Setup the Socket
            _serverSocket = new ClientSocket(ip, 30000);
            _serverSocket.Connected += ServerSocketOnConnected;
            _serverSocket.Disconnected += ServerSocketOnDisconnected;
            _serverSocket.DataReceived += ServerSocketOnDataReceived;
            // -- Setup the buffers
            _inputBuffer = new ByteBuffer();
            _outBuffer = new ByteBuffer();
            _plistBuffer = new ByteBuffer();
            OnlinePlayers = new Dictionary<int, Player>();
            _outBuffer.DataAdded += OutBufferOnDataAdded;
            Battling = false;
            PopulatePackets();
        }

        public void Connect() {
            _serverSocket.Connect();
        }

        public void Disconnect() {
            if (Connected) {
                SendPacket(new ExitPacket());
                _serverSocket.Disconnect("Disconnect called");

                _inputBuffer.GetAllBytes(); // -- Clear the input buffer
                _outBuffer.GetAllBytes(); // -- Clear the output buffer
                _plistBuffer.GetAllBytes(); // -- Clear the plist buffer..
                OnlinePlayers = new Dictionary<int, Player>(); // -- Clear the player list
                Info = new ServerInfo(); // -- Clear the server info
                EncryptionEnabled = false;
                // -- Ready for a reconnect :)
            }
        }

        public void SendChat(string message) {
            SendPacket(new ClientChatPacket {Message = message});
        }

        public void SendInstantMessage(string message, byte playerId) {
            var privatePacket = new PrivateMessagePacket {
                Message = message,
                PlayerId = playerId
            };
            SendPacket(privatePacket);
        }

        public void InterpretPlayerData(byte[] data, bool announce = false) {
            var decompStream = new ByteBuffer();
            decompStream.AddBytes(data);

            while (decompStream.Length != 0) {
                var p = new Player();

                byte[] options = decompStream.ReadByteArray(28);
                var binary = NbMethods.BytesToBinary(options);
                // -- Layout (in bits)
                // -- 8: Player number
                p.Id = Convert.ToInt32(binary.Substring(0, 8), 2);
                binary = binary.Substring(8, binary.Length - 8);
                // -- 3: Game version
                p.GameVersion = Convert.ToInt32(binary.Substring(0, 3), 2);
                binary = binary.Substring(3, binary.Length - 3);
                // -- 4: Picture: 4
                p.Picture = Convert.ToInt32(binary.Substring(0, 4), 2);
                binary = binary.Substring(4, binary.Length - 4);
                // -- 4: GfxVersion: 4
                p.GraphicsVersion = Convert.ToInt32(binary.Substring(0, 4), 2);
                binary = binary.Substring(4, binary.Length - 4);
                // -- 2: Auth: 2,
                p.Authority = Convert.ToInt32(binary.Substring(0, 2), 2);
                binary = binary.Substring(2, binary.Length - 2);
                // -- 1: show team
                var bval = binary.Substring(0, 1);
                p.ShowTeam = Convert.ToBoolean(Convert.ToInt32(bval));
                binary = binary.Substring(1, binary.Length - 1);
                // -- 1: stadium ok
                bval = binary.Substring(0, 1);
                p.StadiumOk = Convert.ToBoolean(Convert.ToInt32(bval));
                binary = binary.Substring(1, binary.Length - 1);
                // -- (for loop, 1-6)
                // -- 9: Pokemon number
                // -- 5: unknown letter
                // -- 1: Shiny?
                // -- (end for)
                p.Team = new List<TeamMember>();
                for (int i = 0; i < 6; i++) {
                    int id = Convert.ToInt32(binary.Substring(0, 9), 2);
                    binary = binary.Substring(9, binary.Length - 9);
                    int unknownLetter = Convert.ToInt32(binary.Substring(0, 5), 2);
                    binary = binary.Substring(5, binary.Length - 5);
                    bval = binary.Substring(0, 1);
                    var shiny = Convert.ToBoolean(Convert.ToInt32(bval));
                    binary = binary.Substring(1, binary.Length - 1);

                    if (!PokemonDatabase.BasePokemonMap.TryGetValue(id, out PokemonSpecies? species))
                        continue; // -- Empty team slot.

                    p.Team.Add(new TeamMember(species) {
                        Shiny = shiny,
                        UnownLetter = (byte)unknownLetter
                    });
                }

                // -- 16: Wins
                p.Wins = Convert.ToInt32(binary.Substring(0, 16), 2);
                binary = binary.Substring(16, binary.Length - 16);
                // -- 16: Losses
                p.Losses = Convert.ToInt32(binary.Substring(0, 16), 2);
                binary = binary.Substring(16, binary.Length - 16);
                // -- 16: Ties
                p.Ties = Convert.ToInt32(binary.Substring(0, 16), 2);
                binary = binary.Substring(16, binary.Length - 16);
                // -- 16: Disconnect
                p.Disconnects = Convert.ToInt32(binary.Substring(0, 16), 2);
                binary = binary.Substring(16, binary.Length - 16);
                // -- 16: Compatibility
                p.Compatibility = Convert.ToInt32(binary.Substring(0, 16), 2);
                binary = binary.Substring(16, binary.Length - 16);
                // -- 16: Rank
                p.TeamPower = Convert.ToInt32(binary.Substring(0, 16), 2);
                binary = binary.Substring(16, binary.Length - 16);
                // -- 11: battleing with
                p.BattlingWith = Convert.ToInt32(binary.Substring(0, 11), 2);
                binary = binary.Substring(11, binary.Length - 11);
                //---- Build string leftovers: (in bytes now)
                // -- 20: name
                string playerName = decompStream.ReadString(20);
                // -- 8: Version
                string playerVersion = decompStream.ReadString(8);
                // -- The Rest; Extra info.
                string playerInfo = decompStream.ReadString(decompStream.Length);
                p.Name = playerName;
                p.Version = playerVersion;
                p.Description = playerInfo;

                OnlinePlayers[p.Id] = p; // -- The server resends a player's data when it changes.

                if (announce)
                    InvokePlayerJoined(p);
            }

            InvokePlayerlistUpdate(OnlinePlayers);
            _plistBuffer = new ByteBuffer();
        }

        public void ParsePlayerList(bool annouce = false) {
            _plistBuffer.ReadByteArray(6); // -- strip the headers..
            byte[] all = _plistBuffer.GetAllBytes();
            byte[] decomp = GZip.Decompress2(all);

            if (decomp == null) {
                Console.WriteLine("Failed to decompress list, stopping.");
                return;
            }
            // -- First byte is a length mark for this entry.

            var decompStream = new ByteBuffer();
            decompStream.AddBytes(decomp);

            while (decompStream.Length > 0) {
                byte length = decompStream.ReadByte();
                InterpretPlayerData(decompStream.ReadByteArray(length));
            }
        }

        public void AddPlayerlistData(byte[] data) {
            _plistBuffer.AddBytes(data);
        }

        #region Client Events

        public event EmptyEventArgs ServerConnected;
        public event EmptyEventArgs ServerDisconnected;
        public event EmptyEventArgs ServerKickedYou;
        public event EmptyEventArgs DuplicateNameKick;
        public event EmptyEventArgs InvalidUserPassword;
        public event EmptyEventArgs IpBanned;
        public event EmptyEventArgs PasswordRequested;
        public event EmptyEventArgs ServerLocked;
        public event EmptyEventArgs ServerFull;

        public event ServerInfoEventArgs ServerInfoReceived;
        public event MessageEventArgs ChatMessageReceived;
        public event MessageEventArgs WelcomeMessageReceived;
        public event MessageEventArgs RegistryMessageReceived;
        public event MessageEventArgs TempBanned;
        public event PrivateMessageEventArgs PmReceived;

        public event PlayerEventArgs PlayerJoined;
        public event PlayerEventArgs PlayerLeft;
        public event PlayerEventArgs PlayerAway;
        public event PlayerEventArgs PlayerBack;
        public event PlayerEventArgs PlayerInfoUpdated;
        public event PlayerEventArgs PlayerKicked;

        public event PlayersEventArgs PlayerlistUpdated;
        public event ChallengeEventArgs ChallengeReceived;

        public void InvokeIpBanned() {
            IpBanned?.Invoke();
        }

        public void InvokeBadPassword() {
            InvalidUserPassword?.Invoke();
        }

        public void InvokeServerLocked() {
            ServerLocked?.Invoke();
        }

        public void InvokeServerFull() {
            ServerFull?.Invoke();
        }
        
        public void InvokeKicked() {
            ServerKickedYou?.Invoke();
            var leavePacket = new ExitPacket();
            SendPacket(leavePacket);
            Disconnect();
        }

        public void InvokeDuplicateName() {
            DuplicateNameKick?.Invoke();
        }
        
        public void InvokePasswordRequested() {
            PasswordRequested?.Invoke();
        }
        
        public void InvokePrivateMessageReceived(Player p, string message) {
            PmReceived?.Invoke(p, message);
        }

        public void InvokePlayerJoined(Player p) {
            PlayerJoined?.Invoke(p);
        }
        public void InvokeServerInfoReceived() {
            ServerInfoReceived?.Invoke(Info);
        }

        public void InvokeChatMessage(string message) {
            ChatMessageReceived?.Invoke(message);
        }
        
        public void InvokeRegistryMessage(string message) {
            RegistryMessageReceived?.Invoke(message);
        }
        
        public void InvokeWelcomeMessage(string message) {
            WelcomeMessageReceived?.Invoke(message);
        }

        public void InvokePlayerlistUpdate(Dictionary<int, Player> players) {
            PlayerlistUpdated?.Invoke(players.Values.ToList());
        }

        public void InvokePlayerUpdate(Player p) {
            PlayerInfoUpdated?.Invoke(p);
        }

        public void InvokePlayerAway(Player p) {
            PlayerAway?.Invoke(p);
        }

        public void InvokePlayerBack(Player p) {
            PlayerBack?.Invoke(p);
        }

        public void InvokePlayerLeft(int id) {
            if (!OnlinePlayers.ContainsKey(id)) {
                return;
            }
            Player p = OnlinePlayers[id];
            PlayerLeft?.Invoke(p);
            OnlinePlayers.Remove(id);
            InvokePlayerlistUpdate(OnlinePlayers);
        }
        #endregion
        #region Socket Events

        public void SendPacket(IPacket packet) {
            lock (_outBuffer) {
                packet.Write(_outBuffer);
            }

        }
        private void OutBufferOnDataAdded() {
            byte[] data = _outBuffer.GetAllBytes();

            if (EncryptionEnabled)
                data = XorModule.XorEncrypt(data);

            var length = (byte)(data.Length - 1); // -- Because netbattle logic..
            var finalData = new[] { length }; // -- Prepend the finished array with the length.
            finalData = finalData.Concat(data).ToArray();
            _serverSocket.Send(finalData);
        }

        private void ServerSocketOnDataReceived(DataReceivedArgs args) {
            lock (_inputBuffer) {
                _inputBuffer.AddBytes(args.Data);
            }

            Handle();
        }

        private void ServerSocketOnDisconnected(SocketDisconnectedArgs args) {
            Connected = false;
            ServerDisconnected?.Invoke();
            Logger.Log(LogType.Debug, $"Disconnected from server {ServerIp}");
        }

        private void ServerSocketOnConnected(SocketConnectedArgs args) {
            Connected = true;
            ServerConnected?.Invoke();
            Logger.Log(LogType.Debug, $"Connected to server {ServerIp}");
        }

        /// <summary>
        /// Populates all valid packets for usage by netbattle.
        /// </summary>
        private void PopulatePackets() {
            _packets = new Dictionary<string, IPacket>();

            // -- Use reflection to load the accepted packet types.
            Type[] types = Assembly.GetAssembly(typeof(IPacket)).GetTypes();
            types = types.Where(t => t.GetInterfaces().Contains(typeof(IPacket))).ToArray();

            foreach (Type type in types) {
                var cmd = (IPacket)Activator.CreateInstance(type);
                _packets.Add(cmd.Command, cmd);
            }
        }

        /// <summary>
        /// Handles incoming data, and waits for data if it is missing.
        /// </summary>
        private void Handle() {
            while (_inputBuffer.Length > 0) {
                lock (_inputBuffer) {
                    int length = _inputBuffer.PeekByte() + 1; // -- Up to 256, so not a byte.

                    if (_inputBuffer.Length - 1 < length)
                        break;

                    _inputBuffer.ReadByte(); // -- Dispose of the length code..

                    byte[] data = EncryptionEnabled
                        ? XorModule.XorDecrypt(_inputBuffer.ReadByteArray(length))
                        : _inputBuffer.ReadByteArray(length);

                    var tempBuffer = new ByteBuffer(); // -- Create a new bytebuffer with the data
                    tempBuffer.AddBytes(data);

                    string cmd = tempBuffer.ReadString(Math.Min(5, tempBuffer.Length)); // -- Get the command to find the regPacket parser..
                    // tempBuffer.ReadByte(); // -- Trim off the colon.

                    // -- try to find the Packet.
                    IPacket packet;
                    if (!_packets.TryGetValue(cmd.Substring(0, Math.Min(4, cmd.Length)), out packet) && !_packets.TryGetValue(cmd, out packet)) {
                        Logger.Log(LogType.Error, "Invalid packet received!!");
                        continue;
                    }

                    try {
                        packet.Read(tempBuffer);
                        packet.Handle(this);
                    }
                    catch (Exception e) when (e is NotImplementedException or EndOfStreamException) {
                        // -- Plenty of packets aren't ported yet, they shouldn't take down the socket thread.
                        Logger.Log(LogType.Warning, $"Couldn't process packet {cmd}: {e.Message}");
                    }
                }
            }
        }
        #endregion

        public void InvokeChallengeReceived(string challengerSlot, BattleModes challengeMode, TerrainTypes challengeTerrain, byte[] rules, int ping) {
            if (!int.TryParse(challengerSlot, out int slot)) {
                Logger.Log(LogType.Error, $"[Weird Server Behavior] Invalid challengerSlot {challengerSlot}");
                return;
            }
            if (!OnlinePlayers.ContainsKey(slot)) {
                Logger.Log(LogType.Error, $"[Weird Server Behavior] Challenge received from unknown slot {slot}");
                return;
            }
            Player challenger = OnlinePlayers[slot];
            challenger.Ping = ping;
            
            ChallengeReceived?.Invoke(challenger, challengeTerrain, challengeMode, rules);
        }

        public void InvokeTempBanned(string banMessage) {
            TempBanned?.Invoke(banMessage);
        }

        public void SendServerPassword(string password) {
            var up = new UserPacket {
                Username = UserSettings.CurrentSettings.Username,
                Password = password,
                ServerPassword =  "potato",
                LoginCode = LoginCode,
                Sid = "WAMABOTBOTEATERABCDEF",
                ClientVersion = "0.9.7",
                ClientPicture = UserSettings.CurrentSettings.IconUsed,
                Options = new bool[10],
                ExtraInfo = UserSettings.CurrentSettings.MoreInfo,
            };

            SendPacket(up);
        }

        public void InvokePlayerKicked(string playerKicked) {
            if (int.TryParse(playerKicked, out int result)) {
                if (OnlinePlayers.TryGetValue(result-1, out Player playerObj)) {
                    PlayerKicked?.Invoke(playerObj);
                    return;
                }
            }
            InvokeChatMessage($"{playerKicked} has been kicked");
        }
    }
}
