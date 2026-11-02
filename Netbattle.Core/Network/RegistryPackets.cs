using Netbattle.Core.Common;

namespace Netbattle.Core.Network {
    public struct ServerRegPacket : IRegPacket {
        public string Command => "SERV";
        public ServerListing ServerData;

        public void Read(ByteBuffer reader) {

            ServerData = new ServerListing {
                ServerNumber = Convert.ToInt32(reader.ReadString(2), 16),
                Name = reader.ReadString(20).Trim(),
                Owner = reader.ReadString(20).Trim(),
                OnlinePlayers = Convert.ToInt32(reader.ReadString(2), 16),
                MaxPlayers = Convert.ToInt32(reader.ReadString(2), 16),
                Ip = reader.ReadString(8),
                Description = reader.ReadString(reader.Length)
            };
            // -- Numbers as sent as HEX strings..
            // -- Description is just the rest of the regPacket :)
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(RegistryClient client) {
            var servData = ServerData;

            if (client.Servers.Any(a => a.ServerNumber == servData.ServerNumber)) {
                client.Servers.Remove(client.Servers.FirstOrDefault(a => a.ServerNumber == servData.ServerNumber));
            }

            client.Servers.Add(ServerData);
            client.RaiseServerListChanged();
        }
    }

    public struct TempbanRegPacket : IRegPacket {
        public string Command => "TBAN";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(RegistryClient client) {
            client.RaiseRegistryBanned();
        }
    }

    public struct DupeRegPacket : IRegPacket {
        public string Command => "MULTI";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(RegistryClient client) {
            client.RaiseDuplicateClient();
        }
    }

    public struct PingRegPacket : IRegPacket {
        public string Command => "PING";
        public void Read(ByteBuffer reader) {
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(RegistryClient client) {
            var pct = new PongRegPacket();
            client.SendPacket(pct);
        }
    }

    public struct PongRegPacket : IRegPacket {
        public string Command => "PONG";
        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            writer.WriteString(Command + ":");
            writer.Purge();
        }

        public void Handle(RegistryClient client) {
            throw new NotImplementedException();
        }
    }

    public struct MoveToTopRegPacket : IRegPacket {
        public string Command => "DISC";

        public void Read(ByteBuffer reader) {
            throw new NotImplementedException();
        }

        public void Write(ByteBuffer writer) {
            throw new NotImplementedException();
        }

        public void Handle(RegistryClient client) {
            throw new NotImplementedException();
        }
    }
}
