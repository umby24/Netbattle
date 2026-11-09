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

        // -- Challenge state, from the VB globals of the same names.
        public int ChallengeNumber { get; set; }          // -- Who the open or pending challenge is with.
        public bool ChallengePending { get; set; }        // -- A challenge has been sent or received and not answered yet.
        public bool ChallengeWindowOpen { get; set; }     // -- VB's MasterServer.ChallengeLoaded.
        public bool ChallengeIsOutgoing { get; set; }     // -- VB's ICalled: we're the challenger.

        public Dictionary<int, Player> OnlinePlayers { get; set; }

        /// <summary>
        /// Your own entry in the player list, which has the server's view of your data, or You before the list arrives.
        /// </summary>
        public Player Me => OnlinePlayers != null && OnlinePlayers.TryGetValue(You.Id, out Player? p) ? p : You;

        public string PlayerName(int id) =>
            OnlinePlayers != null && OnlinePlayers.TryGetValue(id, out Player? p) ? p.Name : $"Player {id}";
        public List<BanListEntry> BanList { get; set; } = new List<BanListEntry>();
        #region Private Variables
        private ClientSocket _serverSocket;
        private ByteBuffer _inputBuffer;
        private ByteBuffer _outBuffer;
        private ByteBuffer _plistBuffer;
        private string _dbModBuffer;
        private Dictionary<string, IPacket> _packets;
        public byte[] LoginCode { get; set; }
        #endregion

        #region Server Info 
        // -- TODO: Make this a class or something.
        public ServerInfo Info;
        #endregion

        public NbClient(string ip) {
            You = new Player {
                Name = UserSettings.CurrentSettings.Username,
                Picture = UserSettings.CurrentSettings.IconUsed,
                Team = UserSettings.CurrentSettings.Team.ToList(),
                Description = UserSettings.CurrentSettings.MoreInfo,
                Version = "0.9.7",
                GameVersion = (int)UserSettings.CurrentSettings.CurrentCompatibilityMode,
                GraphicsVersion = (int)UserSettings.CurrentSettings.CurrentGraphicsMode,
                TeamPower = NbMethods.GetTeamRank(UserSettings.CurrentSettings.Team),
                StationID = NbMethods.DecompressSid(NbMethods.GetSerialNumber())
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
        public void QueryPlayerInfo(string playerName) {
            var queryPacket = new Lookup() {
                PlayerName = playerName
            };
            SendPacket(queryPacket);
        }
        public void QueryPlayerAlias(string playerName) {
            var queryPacket = new Aliases() {
                PlayerName = playerName
            };
            SendPacket(queryPacket);
        }
        
#region Helper Methods
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

        public void AddModData(string data) {
            _dbModBuffer += data;
        }
#endregion
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
        public event EmptyEventArgs BanListUpdated;
        public event UserLookupEventArgs LookupResultReceived;
        public event AliasDataEventArgs AliasDataRecieved;

        public event PlayerEventArgs PlayerJoined;
        public event PlayerEventArgs PlayerLeft;
        public event PlayerEventArgs PlayerAway;
        public event PlayerEventArgs PlayerBack;
        public event PlayerEventArgs PlayerInfoUpdated;
        public event PlayerEventArgs PlayerKickedByServer;
        public event PlayerEventArgs PlayerPingUpdated;
        public event DoublePlayerEventArgs PlayerKicked;
        public event DoublePlayerEventArgs PlayerBanned;
        public event DoublePlayerEventArgs PlayerTempBanned;
        public event DoublePlayerEventArgs BattleStarted;

        public event PlayerEventArgs ChallengeCancelled;
        public event PlayerEventArgs ChallengeAccepted;
        public event PlayerEventArgs ChallengeRefused;
        public event EmptyEventArgs ChallengeEnded;
        public event PlayersEventArgs PlayerlistUpdated;
        public event ChallengeEventArgs ChallengeReceived;
        public event MessageEventArgs WatchBegin;
        public event IntEventArgs WatchRejected;

        public void InvokePingUpdated(Player p)
        {
            PlayerPingUpdated?.Invoke(p);
        }

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

        public void InvokeWatchRejected(int reason) {
            WatchRejected?.Invoke(reason);
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
            if (!args.BaseSocket.IsConnected)
            {
                Logger.Log(LogType.Debug, $"Socket disconnected while receiving data from {ServerIp}");
                return;
            }
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
                    byte[] data;
                    try {
                      data = EncryptionEnabled
                            ? XorModule.XorDecrypt(_inputBuffer.ReadByteArray(length))
                            : _inputBuffer.ReadByteArray(length);
                    }
                    catch (Exception e) {
                        Logger.Log(LogType.Error, $"Error occurred while processing packet: {e.Message}");
                        // -- Data has been read, but we don't know what it is. Just skip it.
                        continue;
                    }

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

        public void SendServerPassword(string password) {
            var up = new UserPacket {
                Username = You.Name,
                Password = password,
                ServerPassword = "potato",
                LoginCode = LoginCode,
                Sid = You.StationID,
                ClientVersion = You.Version,
                ClientPicture = You.Picture,
                Options = new bool[10],
                ExtraInfo = You.Description
            };

            SendPacket(up);
        }

        public void TempBan(string playerName, int duration) {
            var tbp = new TempBan() {
                PlayerName = playerName,
                Duration = duration
            };
            SendPacket(tbp);
        }
        
        public void SidBan(string playerName) {
            var sbp = new SIDBan() {
                PlayerName = playerName
            };
            SendPacket(sbp);
        }
        
        public void IPBan(string playerName) {
            var ibp = new IPBan() {
                PlayerName = playerName
            };
            SendPacket(ibp);
        }

        public void Unban(string tag, string text) {
            var unban = new UnbanPacket() {
                Tag = tag,
                PlayerName = text
            };
            SendPacket(unban);
        }

        internal void ProcessBanList(string data)
        {
            BanList.Clear();
            if (data == "||")
            {
                BanList.Add(new BanListEntry() { Name = "[No current bans]" });
                InvokeBanListUpdate();
                return;
            }
            
            string[] items = data.Split('|');
            if (items.Length > 0)
            {
                string[] tempBans = items[0].Split(',');
                foreach (var t in tempBans)
                {
                    if (string.IsNullOrEmpty(t)) continue;
                    var final = t.Split(':');
                    var entry = new BanListEntry
                    {
                        Name = final[0],
                        Type = "T",
                        Duration = final[1]
                    };
                    BanList.Add(entry);
                }
                string[] sidBans = items[1].Split(',');
                foreach (var s in sidBans)
                {
                    if (string.IsNullOrEmpty(s)) continue;
                    var entry = new BanListEntry
                    {
                        Name = s,
                        Type = "S",
                        Duration = "Permanent"
                    };
                    BanList.Add(entry);
                }
                string[] ipBans = items[2].Split(',');
                foreach (var i in ipBans)
                {
                    if (string.IsNullOrEmpty(i)) continue;
                    var entry = new BanListEntry
                    {
                        Name = i,
                        Type = "I",
                        Duration = "Permanent"
                    };
                    BanList.Add(entry);
                }


            }

            InvokeBanListUpdate();
        }

        public void RequestBanList()
        {
            var blp = new UnbanPacket(); // -- QUIRKY: This is a packet that the server will respond to with the ban list. It doesn't actually unban anyone.
            blp.PlayerName = "";
            blp.Tag = "";
            SendPacket(blp);
        }

        public void InvokeChallengeReceived(int slot, ChallengeSettings settings, string speed) {
            if (!OnlinePlayers.TryGetValue(slot, out Player? challenger)) {
                Logger.Log(LogType.Error, $"[Weird Server Behavior] Challenge received from unknown slot {slot}");
                return;
            }
            if (int.TryParse(speed, out int ping))
                challenger.Ping = ping;

            ChallengeNumber = slot;
            ChallengeIsOutgoing = false;
            ChallengeReceived?.Invoke(challenger, settings);
        }

        public void SendChallenge(Player target, ChallengeSettings settings) {
            SendPacket(new ChallengeReceived {
                PlayerId = target.Id,
                GameVersion = target.GameVersion,
                Settings = settings
            });
            ChallengePending = true;
        }

        public void AcceptChallenge(ChallengeSettings settings) {
            You.BattlingWith = ChallengeNumber;
            SendPacket(new PlayerAcceptedChallenge { PlayerId = ChallengeNumber, Settings = settings });
        }

        public void RefuseChallenge() {
            SendPacket(new PlayerRefusedBattle { ChallengeNumber = (byte)ChallengeNumber });
            ChallengeNumber = 0;
            ChallengePending = false;
        }

        public void CancelChallenge(bool sent) {
            if (sent)
                SendPacket(new ChallengedCancelled { ChallengeNumber = (byte)ChallengeNumber });
            ChallengeNumber = 0;
            ChallengePending = false;
        }

        /// <summary>
        /// The server or the other player ended the challenge: the challenge window should close without replying.
        /// </summary>
        public void InvokeChallengeEnded() {
            ChallengeEnded?.Invoke();
        }

        public void InvokeTempBanned(string banMessage) {
            TempBanned?.Invoke(banMessage);
        }
        public void InvokePlayerKicked(string playerKicked) {
            if (int.TryParse(playerKicked, out int result)) {
                if (OnlinePlayers.TryGetValue(result-1, out Player playerObj)) {
                    PlayerKickedByServer?.Invoke(playerObj);
                    return;
                }
            }
            InvokeChatMessage($"{playerKicked} has been kicked");
        }

        public void InvokePlayerKicked(Player admin, Player kicked) {
            PlayerKicked?.Invoke(admin, kicked);
        }
        
        public void InvokeBanListUpdate()
        {
            BanListUpdated?.Invoke();
        }
        
        internal void InvokeLookupResult(string data)
        {
            if (data.StartsWith('|')) { // -- Server reports no such user.
                LookupResultReceived?.Invoke(new UserLookupResult { Found = false });
                return;
            }

            // -- Split on the first two bars only, so the last login text can't shift the IP.
            var parts = data.Split('|', 3);
            if (parts.Length < 2 || parts[1].Length < 22)
            {
                Logger.Log(LogType.Warning, $"Invalid lookup result received: {data}");
                return;
            }

            var result = new UserLookupResult {
                Found = true,
                Name = parts[0],
                Status = parts[1][0] - '0', // -- Val(Mid(Data, 2, 1))
                StationId = parts[1].Substring(1, 21),
                LastLogin = parts[1].Substring(22),
                IP = parts.Length > 2 ? parts[2] : ""
            };

            LookupResultReceived?.Invoke(result);
        }

        internal void InvokePlayerTempBanned(int playerId, short duration, string playerBanned)
        {
            var gotPlayer = OnlinePlayers.TryGetValue(playerId, out Player banner);
            if (!gotPlayer)
            {
                Logger.Log(LogType.Warning, $"[Weird Server Behavior] PlayerTempBanned invoked for unknown playerId {playerId}");
                return;
            }
            var bannedPlayer = OnlinePlayers.Values.FirstOrDefault(p => string.Equals(p.Name, playerBanned, StringComparison.OrdinalIgnoreCase));
            if (bannedPlayer == null)
            {
                Logger.Log(LogType.Warning, $"[Weird Server Behavior] PlayerTempBanned invoked for unknown playerName {playerBanned}");
                return;
            }

            PlayerTempBanned?.Invoke(banner, bannedPlayer);
        }

        internal void ProcessAliasData(string aliasData)
        {
            List<string> aliases = new List<string>();

            while (aliasData.Length > 0)
            {
                string entry = aliasData.Substring(0, Math.Min(20, aliasData.Length)).Trim();
                if (!string.IsNullOrEmpty(entry))
                    aliases.Add(entry);

                aliasData = aliasData.Substring(Math.Min(20, aliasData.Length));
            }

            InvokeAliasData(aliases);
        }

        private void InvokeAliasData(List<string> aliases)
        {
            AliasDataRecieved?.Invoke(aliases);
        }

        internal void InvokePlayerBanned(byte playerId, string bannedName)
        {
            var gotPlayer = OnlinePlayers.TryGetValue(playerId, out Player banner);
            if (!gotPlayer)
            {
                Logger.Log(LogType.Warning, $"[Weird Server Behavior] PlayerTempBanned invoked for unknown playerId {playerId}");
                return;
            }
            var bannedPlayer = OnlinePlayers.Values.FirstOrDefault(p => string.Equals(p.Name, bannedName, StringComparison.OrdinalIgnoreCase));
            if (bannedPlayer == null)
            {
                Logger.Log(LogType.Warning, $"[Weird Server Behavior] PlayerTempBanned invoked for unknown playerName {bannedPlayer}");
                return;
            }

            PlayerBanned?.Invoke(banner, bannedPlayer);
        }

        internal void InvokeBattleStarted(Player player1, Player player2)
        {
            BattleStarted?.Invoke(player1, player2);
        }

        internal void RaiseChallengeCancelled(Player p)
        {
            ChallengeCancelled?.Invoke(p);
        }

        public void InvokeStartWatching(string watchId) {
            WatchBegin?.Invoke(watchId);
        }

        public void InvokeStartBattle(string battleData) {
            ChallengePending = false;
            ChallengeEnded?.Invoke(); // -- The challenge window closes once the battle starts.
            throw new NotImplementedException("The battle window isn't ported yet.");
        }

        public void InvokeChallengeAccepted(Player challenger) {
            ChallengeAccepted?.Invoke(challenger);
        }

        public void InvokeChallengeRefused(Player challenger) {
            ChallengeRefused?.Invoke(challenger);
        }

        public void InvokeMaxUsersUpdate() {
            throw new NotImplementedException();
        }

        public void InvokeServerShutdown() {
            throw new NotImplementedException();
        }

        public void ActivateDbMod() {
            throw new NotImplementedException();
        }
    }
}
