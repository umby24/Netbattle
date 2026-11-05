using Sockets;
using Sockets.EventArgs;

namespace Netbattle.Core.Network
{
    public class NbServer
    {
        public bool UseXOR { get; set; }
        
        private ServerSocket _serverSocket;

        public NbServer()
        {
            _serverSocket = new ServerSocket(30000);
            _serverSocket.IncomingClient += ServerSocketOnIncomingClient;
            try
            {
                _serverSocket.Listen();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            
        }

        private void ServerSocketOnIncomingClient(IncomingEventArgs args)
        {
            throw new NotImplementedException();
        }
    }
}