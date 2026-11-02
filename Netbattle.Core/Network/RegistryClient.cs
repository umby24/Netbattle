using Netbattle.Core.Common;
using Netbattle.Core.Database;
using Sockets;
using Sockets.EventArgs;

namespace Netbattle.Core.Network;

public class RegistryClient {
    private const int RegPortClient = 30002;
    private const int RegPortServer = 30001;
    private readonly ClientSocket _regSock;
    private readonly ByteBuffer _sendBuffer;
    private readonly ByteBuffer _receiveBuffer;
    private readonly bool _canReceive;
    private Dictionary<string, IRegPacket> _packets;
    public List<ServerListing> Servers;

    public RegistryClient(bool serverMode = false) {
        _sendBuffer = new ByteBuffer();
        _receiveBuffer = new ByteBuffer();
        _sendBuffer.DataAdded += SendBufferOnDataAdded;
        Servers = new List<ServerListing>();
        PopulatePackets();

        if (!serverMode) {
            _regSock = new ClientSocket(UserSettings.CurrentSettings.RegistryAddress, RegPortClient);
        }
        else {
            _regSock = new ClientSocket(UserSettings.CurrentSettings.RegistryAddress, RegPortServer);
        }

        _regSock.DataReceived += ReceivedRegistryData;
        _regSock.Disconnected += RegistryDisconnected;
        _canReceive = true;

        _regSock.Connect();
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
            
            // -- Every packet is prefixed with a length byte, so we can check if we have enough data to read the packet.
            var length = (byte) (_receiveBuffer.PeekByte() + 1);

            if (_receiveBuffer.Length - 1 < length)
                return; // -- Not enough yet, wait for more data to come in.

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
                Logger.Log(LogType.Error,$"[RegistryClient] Unknown packet received: {cmd}");
                return;
            }

            regPacket.Read(tempBuffer);
            regPacket.Handle(this);
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

    public void SendPacket(IRegPacket regPacket) {
        lock (_sendBuffer) {
            regPacket.Write(_sendBuffer);
        }
    }

    private void RegistryDisconnected(SocketDisconnectedArgs args) {
        RaiseDisconnected();
    }

    private void ReceivedRegistryData(DataReceivedArgs args) {
        if (!_canReceive)
            return;

        lock (_receiveBuffer) {
            _receiveBuffer.AddBytes(args.Data);
        }

        HandlePackets();
    }
    
    public event NoArgs DuplicateClient;
    public event NoArgs RegistryBanned;
    public event NoArgs ServerListChanged;
    public event NoArgs Disconnected;
    
    public void RaiseDuplicateClient() {
        DuplicateClient?.Invoke();
    }

    public void RaiseRegistryBanned() {
        RegistryBanned?.Invoke();
    }

    public void RaiseServerListChanged() {
        ServerListChanged?.Invoke();
    }

    public void RaiseDisconnected() {
        Disconnected?.Invoke();
    }
}