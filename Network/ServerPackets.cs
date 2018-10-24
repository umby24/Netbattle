using System;
using System.Security.Cryptography;
using System.Text;
using Netbattle.Common;
using Netbattle.Database;

namespace Netbattle.Network {
    #region Server-Bound
    public struct ClientChatPacket : IPacket {
        public string Command => "CHAT";
        public string Message { get; set; }

        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            writer.WriteString(Command + ":");
            writer.WriteString(Message);
            writer.Purge();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct UserPacket : IPacket {
        public string Command => "USER";
        public string Password; // -- Optional, 16 chrs.
        public string Username; // -- 20 chrs
        public string ClientVersion; // -- 7 chrs.
        public string Sid; // -- dunno how many.. 21?
        public string ServerPassword; // -- 15 characters, blank if not included, aka your saved password MD5, hexed, and then turned into a chr.
        // -- Build & Chr$(Val("&H" & Mid(Temp, X, 2)))
        // -- 10 bytes for various client options, then turned into a bitstring and later, bytes.
        public bool[] Options;
        public int ClientPicture; // -- turned into a hex string and included..
        public bool AllowViewing; // -- ???
        public byte[] LoginCode; // -- Code sent by server, to be returned by the client.
        public string ExtraInfo;

        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            writer.WriteString(Command + ":");

            if (!string.IsNullOrEmpty(Password))
                writer.WriteString(Password);

            writer.WriteString(Username.PadRight(20));
            writer.WriteString(ClientVersion.PadRight(7));
            writer.AddBytes(NbMethods.CompressSid(Sid));

            MD5 md5Er = MD5.Create();
            byte[] bytes = md5Er.ComputeHash(Encoding.ASCII.GetBytes(ServerPassword));
            writer.AddBytes(bytes);

            string stuffBytes = "30 42 31 38".Replace(" ", "");
            var temp = "";

            foreach (bool option in Options) {
                temp += option ? "1" : "0";
            }

            writer.AddBytes(NbMethods.Bin2Chr(temp));
            writer.WriteString(NbMethods.FixedHex(ClientPicture, 2));
            writer.AddBytes(NbMethods.StringToByteArray(stuffBytes));
            writer.AddBytes(LoginCode);
            writer.WriteString(ExtraInfo);
            // -- we'll just end there..
            writer.Purge();

        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// SERVERBOUND ONLY
    /// </summary>
    public struct ExitPacket : IPacket {
        public string Command => "EXIT";

        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            writer.WriteString(Command + ":");
            writer.Purge();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }


    public struct Team : IPacket {
        public string Command => "TEAM";

        public const string StaticTeam =
            "54 45 41 4D 3A 30 50 37 44 75 67 74 72 69 6F 20 20 20 20 20 20 20 20 19 BC 85 07 00 03 75 3A F1 81 FF FB DE F0 27 E0 07 E0 00 00 41 65 72 6F 64 61 63 74 79 6C 20 20 20 20 20 47 3C 85 08 00 0F C1 B9 6D 4F FF FB FF F0 27 E0 07 E0 00 00 41 6C 61 6B 61 7A 61 6D 20 20 20 20 20 20 20 20 BC 8A 54 00 06 34 C2 85 45 FF FF FF F8 00 00 07 E7 E0 20 47 65 6E 67 61 72 20 20 20 20 20 20 20 20 20 2F 3C 85 54 00 05 03 1A 61 D1 FF FF FF F8 00 00 07 E7 E0 00 45 6C 65 63 74 72 6F 64 65 20 20 20 20 20 20 32 BC 86 1E 00 10 B1 E3 9D D1 FF FF FF F8 00 20 07 E7 E0 00 53 74 61 72 6D 69 65 20 20 20 20 20 20 20 20 3C BC 85 69 00 10 C4 C3 65 D1 FF FF FF F8 00 04 07 E0 04 00";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            byte[] data = NbMethods.StringToByteArray(StaticTeam.Replace(" ", ""));
            writer.AddBytes(data);
            writer.Purge();
        }

        public void Handle(NbClient client) {

        }
    }
    #endregion

    public struct RequestServerPassword : IPacket {
        public string Command => "RPWD";

        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct PasswordRefused : IPacket {
        public string Command => "PWDR";

        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct RequestName : IPacket {
        public string Command => "REQN";

        public bool UseEncryption;
        public string ServerVersion;
        public int MaxUsers;
        public int OnlineUsers;
        public int FloodTolerance;
        public string ServerName;
        public byte[] SigninCode;

        public void Read(ByteBuffer reader) {
            var useEnc = (char)reader.ReadByte();

            if (useEnc == '1')
                UseEncryption = true;

            ServerVersion = reader.ReadString(10);
            MaxUsers = Convert.ToInt32(reader.ReadString(3), 16);
            FloodTolerance = Convert.ToInt32(reader.ReadString(2), 16);
            OnlineUsers = Convert.ToInt32(reader.ReadString(3), 16);
            ServerName = reader.ReadString(20);
            SigninCode = reader.ReadByteArray(5);
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            // -- Display some stuff, enable xor if needed, set server code, and SEND
            // -- USER packet.
            client.EncryptionEnabled = UseEncryption;
            var info = new ServerInfo {
                FloodTolerance = FloodTolerance,
                MaxPlayers = MaxUsers,
                OnlinePlayers = OnlineUsers,
                ServerName =  ServerName,
                ServerVersion = ServerVersion
            };

            client.Info = info;
            client.InvokeServerInfoReceived();
             // -- Reply..
            var up = new UserPacket {
                Username = UserSettings.CurrentSettings.Username,
                ServerPassword = "potato",
                Sid = "WAMABOTBOTEATERABCDEF",
                ClientVersion = "0.9.7",
                LoginCode = SigninCode,
                ClientPicture = UserSettings.CurrentSettings.IconUsed,
                Options = new bool[10],
                ExtraInfo = UserSettings.CurrentSettings.MoreInfo,
                
            };

            client.SendPacket(up);
        }
    }

    public struct NameRefused : IPacket {
        public string Command => "NAMR";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct IpBanned : IPacket {
        public string Command => "BANU";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct TempBanned : IPacket {
        public string Command => "TBAN";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct ServerLocked : IPacket {
        public string Command => "NNPL";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct ServerFull : IPacket {
        public string Command => "BUSY";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct RequestTeam : IPacket {
        public string Command => "RQTM";
        public int YourPid;
        public int Authority;
        public int Wins;
        public int Losses;
        public int Ties;
        public int Disconnects;

        public void Read(ByteBuffer reader) {
            string data = Encoding.ASCII.GetString(reader.GetAllBytes());
            string[] opts = data.Split(',');
            YourPid = int.Parse(opts[0]);
            Authority = int.Parse(opts[1]);
            Wins = int.Parse(opts[2]);
            Losses = int.Parse(opts[3]);
            Ties = int.Parse(opts[4]);
            Disconnects = int.Parse(opts[5]);
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            client.You.Wins = Wins;
            client.You.Losses = Losses;
            client.You.Ties = Ties;
            client.You.Disconnects = Disconnects;
            client.You.Id = YourPid;
            client.You.Authority = Authority;

            var tm = new Team();
            client.SendPacket(tm);
        }
    }


    // -- Gziped, parted, list of players..
    public struct PlayerList : IPacket {
        public string Command => "/WHO";
        public byte[] ListData;
        // -- First two bytes tell how many packets this is chunked into..
        // -- The rest is a compressed, weird format of data.. lel..

        public void Read(ByteBuffer reader) {
            ListData = reader.GetAllBytes();

        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            client.AddPlayerlistData(ListData);

            if (ListData.Length < 200)
                client.ParsePlayerList();
        }
    }

    public struct DatabaseMod : IPacket {
        public string Command => "DBMD";
        public void Read(ByteBuffer reader) {
        //    throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
        //    throw new NotImplementedException();
        }
    }

    public struct RequestVersion : IPacket {
        public string Command => "REQV";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct RequestPicture : IPacket {
        public string Command => "REQP";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct RequestUserPassword : IPacket {
        public string Command => "RUSP";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct UserPasswordRefused : IPacket {
        public string Command => "UPWR";
        public void Read(ByteBuffer reader) {
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct MultiConnected : IPacket {
        public string Command => "NOIP";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct AuthChange : IPacket {
        public string Command => "AUTH";
        public int PlayerId { get; set; }
        public int AuthLevel { get; set; }

        public void Read(ByteBuffer reader) {
            PlayerId = Convert.ToInt32(reader.ReadString(3)); // -- Player to update
            AuthLevel = Convert.ToInt32(reader.ReadString(reader.Length)); // -- Their new auth.
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            if (!client.OnlinePlayers.ContainsKey(PlayerId))
                return;

            client.OnlinePlayers[PlayerId].Authority = AuthLevel;
            client.InvokePlayerUpdate(client.OnlinePlayers[PlayerId]);
        }
    }

    public struct YourNumber : IPacket {
        public string Command => "YNUM";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct PlayerList2 : IPacket {
        public string Command => "PLST";
        

        public void Read(ByteBuffer reader) {
           
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            
        }
    }

    public struct PlayerInfo : IPacket {
        public string Command => "PNFO";
        public byte[] ListData;

        public void Read(ByteBuffer reader) {
            ListData = reader.ReadByteArray(reader.Length);
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            client.InterpretPlayerData(ListData, true);
        }
    }

    public struct PlayerUpdate : IPacket {
        public string Command => "PUPD";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct NewPlayer : IPacket {
        public string Command => "NPLY";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct PlayerDisconnect : IPacket {
        public string Command => "PDIS";
        public int PlayerId { get; set; }

        public void Read(ByteBuffer reader) {
            PlayerId = Convert.ToInt32(((char)reader.ReadByte()).ToString());
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            client.InvokePlayerLeft(PlayerId);
        }
    }

    public struct PlayerBack : IPacket {
        public string Command => "BACK";
        public int PlayerId { get; set; }

        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct RequestUpdate : IPacket { // -- Guessing here, no comments in og source.
        public string Command => "RPUD";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct ChatMessage : IPacket {
        public string Command => "CMSG";
        public string Message;

        public void Read(ByteBuffer reader) {
            Message = reader.ReadString(reader.Length);
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            client.InvokeChatMessage(Message);
        }
    }

    public struct PrivateMessagePacket : IPacket {
        public string Command => "IMCH";
        public byte PlayerId { get; set; }
        public string Message { get; set; }

        public void Read(ByteBuffer reader) {
            PlayerId = reader.ReadByte();
            Message = reader.ReadString(reader.Length);
        }

        public void Write(ByteBuffer writer) {
            writer.WriteString(Command + ":");
            writer.WriteByte(PlayerId);
            writer.WriteString(Message);
            writer.Purge();
        }

        public void Handle(NbClient client) {
            Player thePlayer;

            if (!client.OnlinePlayers.TryGetValue((int) PlayerId, out thePlayer)) {
                return;
            }

            client.InvokePrivateMessageReceived(thePlayer, Message);
        }
    }

    public struct ServerMessage : IPacket {
        public string Command => "SMSG";
        public string Message;

        public void Read(ByteBuffer reader) {
            Message = reader.ReadString(reader.Length);
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            client.InvokeWelcomeMessage(Message);
        }
    }

    public struct ServerQuit : IPacket {
        public string Command => "SVRQU";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct PlayerKicked : IPacket {
        public string Command => "KICK";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct KickYou : IPacket {
        public string Command => "KCKU";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct BanYou : IPacket {
        public string Command => "ILLM";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct MaxUserChange : IPacket {
        public string Command => "MUCG";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct ChallengeReceived : IPacket {
        public string Command => "CHLN";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct PlayerBusy : IPacket {
        public string Command => "PBSY";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct PlayerInBattle : IPacket {
        public string Command => "SBSY";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct PlayerRefusedBattle : IPacket {
        public string Command => "PREF";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct ServerRefusedBattle : IPacket {
        public string Command => "SREF";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct PlayerAcceptedChallenge : IPacket {
        public string Command => "PACC";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct StartBattle : IPacket {
        public string Command => "SBAT";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct StartWatch : IPacket {
        public string Command => "SWAT";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct WorldRefused : IPacket {
        public string Command => "WREF";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }

// -- Battle refused for some pre-defined reason.
    }

    public struct ChallengedCancelled : IPacket {
        public string Command => "PCAN";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct SilentKick : IPacket {
        public string Command => "BOOT";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct Ping : IPacket {
        public string Command => "PING";

        public void Read(ByteBuffer reader) {
            
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            var po = new Pong();
            client.SendPacket(po);
        }
    }

    public struct Pong : IPacket {
        public string Command => "PONG";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            writer.WriteString(Command + ":");
            writer.Purge();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct FloodcountChange : IPacket {
        public string Command => "FTCG";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct PlayersBusy : IPacket {
        public string Command => "PRBS";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct PlayerAway : IPacket {
        public string Command => "AWAY";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct PlayerPing : IPacket {
        public string Command => "PSPD";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }

// -- 'Player Speed', lel.
    }

    public struct ModKick : IPacket {
        public string Command => "MKCK";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct ModBan : IPacket {
        public string Command => "MBAN";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct BanResult : IPacket {
        public string Command => "BRLT";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct Lookup : IPacket {
        public string Command => "LOOK";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct Aliases : IPacket {
        public string Command => "ALIA";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct BanLookup : IPacket {
        public string Command => "BANL";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct ModTempBan : IPacket {
        public string Command => "MTBN";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            throw new NotImplementedException();
        }
    }

    public struct RegistryMessage : IPacket {
        public string Command => "MASS";
        public string Message { get; set; }

        public void Read(ByteBuffer reader) {
            Message = reader.ReadString(reader.Length);
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(NbClient client) {
            client.InvokeChatMessage(Message);
        }
    }

}
