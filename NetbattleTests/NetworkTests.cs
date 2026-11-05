using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Netbattle.Common;
using Netbattle.Core.Common;
using Netbattle.Core.Network;
using Netbattle.Network;

namespace NetbattleTests {
    [TestClass]
    public class ByteBufferTests {
        [TestMethod]
        public void IntegersAreBigEndian() {
            var buffer = new ByteBuffer();
            buffer.WriteInt(0x01020304);

            CollectionAssert.AreEqual(new byte[] { 1, 2, 3, 4 }, buffer.GetAllBytes());
        }

        [TestMethod]
        public void ReadsConsumeTheBuffer() {
            var buffer = new ByteBuffer();
            buffer.AddBytes(new byte[] { 7, 0, 1, 0x41, 0x42 });

            Assert.AreEqual(7, buffer.ReadByte());
            Assert.AreEqual(1, buffer.ReadShort());
            Assert.AreEqual("AB", buffer.ReadString(2));
            Assert.AreEqual(0, buffer.Length);
        }

        [TestMethod]
        public void ReadingPastTheEndThrowsInsteadOfHanging() {
            var buffer = new ByteBuffer();
            buffer.AddBytes(new byte[] { 1 });

            var read = Task.Run(() => buffer.ReadString(5));
            Assert.AreEqual(0, Task.WaitAny(new Task[] { read }, TimeSpan.FromSeconds(2)), "ReadString spun forever.");
            Assert.IsInstanceOfType(read.Exception?.InnerException, typeof(EndOfStreamException));
        }
    }

    [TestClass]
    public class XorModuleTests {
        [DataTestMethod]
        [DataRow("CHAT:Hello")]
        [DataRow("CHAT:Hello!")]
        [DataRow("PONG:")]
        public void RoundTrips(string message) {
            for (var i = 0; i < 200; i++) { // -- The key is random, so try plenty.
                byte[] encrypted = XorModule.XorEncrypt(Encoding.ASCII.GetBytes(message));
                Assert.AreEqual(message, Encoding.ASCII.GetString(XorModule.XorDecrypt(encrypted)));
            }
        }

        [TestMethod]
        public void AddsEightTrailerBytes() {
            Assert.AreEqual(16, XorModule.XorEncrypt(new byte[8]).Length);
        }

        [TestMethod]
        public void TamperedPacketFailsChecksum() {
            byte[] encrypted = XorModule.XorEncrypt(Encoding.ASCII.GetBytes("CHAT:Hello!!"));
            encrypted[3] ^= 0xFF;

            Assert.ThrowsException<Exception>(() => XorModule.XorDecrypt(encrypted));
        }
    }

    [TestClass]
    public class PlayerDataTests {
        private static byte[] BuildPlayer(int id, int firstSpecies, int wins, int losses, int ties, int disconnects, int rank, int battlingWith) {
            var bits = new StringBuilder();
            bits.Append(NbMethods.Dec2Bin(id, 8));
            bits.Append(NbMethods.Dec2Bin(3, 3)); // -- Game version
            bits.Append(NbMethods.Dec2Bin(5, 4)); // -- Picture
            bits.Append(NbMethods.Dec2Bin(8, 4)); // -- Graphics version
            bits.Append(NbMethods.Dec2Bin(1, 2)); // -- Authority
            bits.Append("11"); // -- Show team, stadium ok

            for (var i = 0; i < 6; i++) {
                bits.Append(NbMethods.Dec2Bin(i == 0 ? firstSpecies : 0, 9));
                bits.Append(NbMethods.Dec2Bin(0, 5));
                bits.Append(i == 0 ? "1" : "0");
            }

            bits.Append(NbMethods.Dec2Bin(wins, 16));
            bits.Append(NbMethods.Dec2Bin(losses, 16));
            bits.Append(NbMethods.Dec2Bin(ties, 16));
            bits.Append(NbMethods.Dec2Bin(disconnects, 16));
            bits.Append(NbMethods.Dec2Bin(0, 16)); // -- Compatibility
            bits.Append(NbMethods.Dec2Bin(rank, 16));
            bits.Append(NbMethods.Dec2Bin(battlingWith, 11));

            byte[] packed = NbMethods.BinaryToBytes(bits.ToString());
            byte[] text = Encoding.ASCII.GetBytes("Trainer".PadRight(20) + "0.9.7".PadRight(8) + "About me");
            return packed.Concat(text).ToArray();
        }

        private static NbClient MakeClient() => new NbClient("127.0.0.1");

        [TestMethod]
        public void InterpretPlayerData_ReadsVbBitLayout() {
            NbClient client = MakeClient();
            client.InterpretPlayerData(BuildPlayer(12, TestSetup.Pikachu, 1234, 567, 89, 10, 4321, 1025));

            Player player = client.OnlinePlayers[12];
            Assert.AreEqual(1234, player.Wins);
            Assert.AreEqual(567, player.Losses);
            Assert.AreEqual(89, player.Ties);
            Assert.AreEqual(10, player.Disconnects);
            Assert.AreEqual(4321, player.TeamPower);
            Assert.AreEqual(1025, player.BattlingWith);
            Assert.AreEqual("Trainer", player.Name);
            Assert.AreEqual("0.9.7", player.Version);
            Assert.AreEqual("About me", player.Description);
        }

        [TestMethod]
        public void InterpretPlayerData_SkipsEmptyTeamSlots() {
            NbClient client = MakeClient();
            client.InterpretPlayerData(BuildPlayer(2, TestSetup.Pikachu, 0, 0, 0, 0, 0, 0));

            Assert.AreEqual(1, client.OnlinePlayers[2].Team.Count);
            Assert.IsTrue(client.OnlinePlayers[2].Team[0].Shiny);
        }

        [TestMethod]
        public void InterpretPlayerData_UpdatesKnownPlayer() {
            NbClient client = MakeClient();
            client.InterpretPlayerData(BuildPlayer(2, TestSetup.Pikachu, 1, 0, 0, 0, 0, 0));
            client.InterpretPlayerData(BuildPlayer(2, TestSetup.Pikachu, 2, 0, 0, 0, 0, 0));

            Assert.AreEqual(2, client.OnlinePlayers[2].Wins);
        }

        [DataTestMethod]
        [DataRow("7", 7)]
        [DataRow("123", 123)]
        public void PlayerAway_ParsesWholeNumber(string data, int expected) { // -- MasterServer.frm: Val(Data)
            var buffer = new ByteBuffer();
            buffer.AddBytes(Encoding.ASCII.GetBytes(data));
            var packet = new PlayerAway();
            packet.Read(buffer);

            Assert.AreEqual(expected, packet.Socket);
        }

        [TestMethod]
        public void PlayerBack_ParsesWholeNumber() {
            var buffer = new ByteBuffer();
            buffer.AddBytes(Encoding.ASCII.GetBytes("123"));
            var packet = new PlayerBack();
            packet.Read(buffer);

            Assert.AreEqual(123, packet.PlayerId);
        }

        [TestMethod]
        public void PlayerDisconnect_ParsesWholeNumber() {
            var buffer = new ByteBuffer();
            buffer.AddBytes(Encoding.ASCII.GetBytes("123"));
            var packet = new PlayerDisconnect();
            packet.Read(buffer);

            Assert.AreEqual(123, packet.PlayerId);
        }

        [TestMethod]
        public void PlayerAway_SetsVbAwayMarker() {
            NbClient client = MakeClient();
            client.InterpretPlayerData(BuildPlayer(2, TestSetup.Pikachu, 0, 0, 0, 0, 0, 0));

            new PlayerAway { Socket = 2 }.Handle(client);
            Assert.AreEqual(1025, client.OnlinePlayers[2].BattlingWith);

            new PlayerBack { PlayerId = 2 }.Handle(client);
            Assert.AreEqual(0, client.OnlinePlayers[2].BattlingWith);
        }

        [TestMethod]
        public void PlayerAway_IgnoresUnknownPlayer() {
            new PlayerAway { Socket = 99 }.Handle(MakeClient());
            new PlayerBack { PlayerId = 99 }.Handle(MakeClient());
        }

        [TestMethod]
        public void TeamPacket_WritesGraphicsAsOneHexDigit() { // -- MasterServer.frm PrepareTeamInfo
            var buffer = new ByteBuffer();
            var team = new TeamMember[6];
            team[0] = TestSetup.MakeAdvancePokemon();

            new Team { AllowViewing = true, TeamPower = 100, graphics = GraphicsMode.nbGFXSml, Pokemon = team }.Write(buffer);
            byte[] written = buffer.GetAllBytes();

            Assert.AreEqual(5 + 3 + 6 * 35, written.Length);
            Assert.AreEqual((byte)'A', written[7]);
        }

        [TestMethod]
        public void RegistryPackets_ReadAllTheirData() {
            var buffer = new ByteBuffer();
            buffer.AddBytes(Encoding.ASCII.GetBytes("0"));

            new TempbanRegPacket().Read(buffer);
            new DupeRegPacket().Read(buffer);
            new MoveToTopRegPacket().Read(buffer);
        }
    }

    [TestClass]
    public class DatabaseTests {
        [TestMethod]
        public void Species_LoadsAllEntries() {
            Assert.IsTrue(PokemonDatabase.BasePokemon.Count >= 386, $"{PokemonDatabase.BasePokemon.Count} species");
            Assert.AreEqual("Pikachu", PokemonDatabase.BasePokemonMap[TestSetup.Pikachu].Name);
            Assert.AreEqual(35, PokemonDatabase.BasePokemonMap[TestSetup.Pikachu].BaseStats[StatTypes.nbHP]);
        }

        [DataTestMethod]
        [DataRow(Elements.nbFire, Elements.nbGrass, 2f)]
        [DataRow(Elements.nbWater, Elements.nbFire, 2f)]
        [DataRow(Elements.nbElectr, Elements.nbGround, 0f)]
        [DataRow(Elements.nbNormal, Elements.nbGhost, 0f)]
        [DataRow(Elements.nbFire, Elements.nbSteel, 2f)]
        [DataRow(Elements.nbFight, Elements.nbNormal, 2f)]
        public void TypeMatrix_MatchesTypeChart(Elements attacker, Elements defender, float expected) {
            Assert.AreEqual(expected, Netbattle.Core.Database.TypeDatabase.BattleMatrix[(int)attacker, (int)defender]);
        }
    }
}
