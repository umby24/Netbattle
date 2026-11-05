using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Netbattle.Common;
using Netbattle.Core.Common;
using Netbattle.Core.Database;
using Netbattle.Core.Files;
using Netbattle.Core.Pokemon;

namespace NetbattleTests {
    [TestClass]
    public class SaveFileTests {
        private string _directory;

        [TestInitialize]
        public void MakeDirectory() {
            _directory = Path.Combine(Path.GetTempPath(), "NetbattleTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_directory);
        }

        [TestCleanup]
        public void RemoveDirectory() {
            Directory.Delete(_directory, true);
        }

        private string PathTo(string name) => Path.Combine(_directory, name);

        private static void SetSettingsWith(TeamMember member) {
            UserSettings.CurrentSettings = new UserSettings {
                CurrentCompatibilityMode = CompatModes.nbFullAdvance,
                CurrentGraphicsMode = GraphicsMode.nbGFXEme,
                Username = "Tester",
                IconUsed = 1,
                RegistryAddress = "registry.test",
                Team = new[] { member, null, null, null, null, null }
            };
        }

        private static void AssertSameStats(TeamMember expected, TeamMember actual) {
            foreach (StatTypes stat in StatSet.All) {
                Assert.AreEqual(expected.DVs[stat], actual.DVs[stat], $"DV {stat}");
                Assert.AreEqual(expected.EVs[stat], actual.EVs[stat], $"EV {stat}");
                Assert.AreEqual(expected.CurrentStats[stat], actual.CurrentStats[stat], $"Stat {stat}");
            }
        }

        [TestMethod]
        public void Jnb_RoundTripKeepsStats() { // -- The startup bug: team.jnb came back with zeroed stats.
            TeamMember member = TestSetup.MakeAdvancePokemon();
            SetSettingsWith(member);
            string path = PathTo("team.jnb");

            UserSettings.Save(path);
            UserSettings.Load(path);

            TeamMember loaded = UserSettings.CurrentSettings.Team[0];
            AssertSameStats(member, loaded);
            Assert.AreEqual(member.Item, loaded.Item);
            Assert.AreEqual(member.NatureNum, loaded.NatureNum);
            Assert.AreEqual(4, loaded.Move.Length);
        }

        [TestMethod]
        public void Jnb_LoadsMatchFreshPnbLoad() { // -- What you noticed: both load paths must show the same stats.
            TeamMember member = TestSetup.MakeAdvancePokemon();
            SetSettingsWith(member);
            string jnbPath = PathTo("team.jnb");
            string pnbPath = PathTo("team.pnb");
            WritePnb(pnbPath, member);

            UserSettings.Save(jnbPath);
            UserSettings.Load(jnbPath);
            var pnb = new PnbFile(pnbPath);
            pnb.Load();

            AssertSameStats(pnb.Team[0], UserSettings.CurrentSettings.Team[0]);
        }

        [TestMethod]
        public void Jnb_LinksToTheSharedSpecies() {
            SetSettingsWith(TestSetup.MakeAdvancePokemon());
            string path = PathTo("team.jnb");

            UserSettings.Save(path);
            UserSettings.Load(path);

            Assert.AreSame(PokemonDatabase.BasePokemonMap[TestSetup.Pikachu], UserSettings.CurrentSettings.Team[0].Pokemon);
        }

        [TestMethod]
        public void Jnb_DoesNotStoreTheSpeciesData() {
            SetSettingsWith(TestSetup.MakeAdvancePokemon());
            string path = PathTo("team.jnb");

            UserSettings.Save(path);

            Assert.IsTrue(new FileInfo(path).Length < 10_000, $"team.jnb was {new FileInfo(path).Length} bytes");
        }

        [TestMethod]
        public void Jnb_ReadsOldFilesWithEmbeddedSpecies() {
            string path = PathTo("old.jnb");
            File.WriteAllText(path, """
                {
                  "CurrentCompatibilityMode": 3,
                  "Team": [
                    { "Nickname": "Old", "Level": 100, "Pokemon": { "No": 25, "BaseStats": { "Total": 0 } },
                      "DVs": { "Total": 0 }, "EVs": { "Total": 0 }, "Move": [1, 2, 3, 4] },
                    null, null, null, null, null
                  ]
                }
                """);

            UserSettings.Load(path);

            TeamMember loaded = UserSettings.CurrentSettings.Team[0];
            Assert.AreEqual("Pikachu", loaded.Pokemon.Name);
            Assert.AreEqual(35, loaded.Pokemon.BaseStats[StatTypes.nbHP]);
        }

        [TestMethod]
        public void Jnb_MissingFileFallsBackToNewSettings() {
            UserSettings.Load(PathTo("missing.jnb"));

            Assert.IsNotNull(UserSettings.CurrentSettings);
            Assert.AreEqual(6, UserSettings.CurrentSettings.Team.Length);
        }

        [TestMethod]
        public void Load_ReadsPnbPath() {
            string path = PathTo("team.pnb");
            WritePnb(path, TestSetup.MakeAdvancePokemon());

            UserSettings.Load(null, path);

            Assert.AreEqual("Tester", UserSettings.CurrentSettings.Username);
            Assert.AreEqual(CompatModes.nbFullAdvance, UserSettings.CurrentSettings.CurrentCompatibilityMode);
            Assert.AreEqual(TestSetup.Pikachu, UserSettings.CurrentSettings.Team[0].Pokemon.No);
        }

        private static void WritePnb(string path, TeamMember member) {
            var pnb = new PnbFile(path) {
                Name = "Tester",
                ExtraInfo = "Extra",
                WinMessage = "Won",
                LoseMessage = "Lost",
                TBMode = 3,
                CurrentPicture = 2,
                GraphicsVersion = (byte)GraphicsMode.nbGFXEme,
                Team = new[] { member, null, null, null, null, null }
            };
            pnb.Save();
        }

        [TestMethod]
        public void Pnb_HeaderMatchesVbFormat() { // -- TeamBuilder.frm: " PNB4.1" then Chr$(Len(.Name)) & .Name
            string path = PathTo("team.pnb");
            WritePnb(path, TestSetup.MakeAdvancePokemon());

            byte[] file = File.ReadAllBytes(path);

            Assert.AreEqual(" PNB4.1", Encoding.ASCII.GetString(file, 0, 7));
            Assert.AreEqual(6, file[7]);
            Assert.AreEqual("Tester", Encoding.ASCII.GetString(file, 8, 6));
        }

        [TestMethod]
        public void Pnb_RoundTripsTeamAndText() {
            string path = PathTo("team.pnb");
            TeamMember member = TestSetup.MakeAdvancePokemon();
            WritePnb(path, member);

            var loaded = new PnbFile(path);
            loaded.Load();

            Assert.AreEqual("Tester", loaded.Name);
            Assert.AreEqual("Lost", loaded.LoseMessage);
            Assert.AreEqual(2, loaded.CurrentPicture);
            Assert.AreEqual("Testmon", loaded.Team[0].Nickname);
            Assert.AreEqual(Items.nbLeftovers, loaded.Team[0].Item);
            Assert.IsNull(loaded.Team[1]);
            AssertSameStats(member, loaded.Team[0]);
        }

        [TestMethod]
        public void Pnb_LoadAppliesNatureToStats() {
            string path = PathTo("team.pnb");
            WritePnb(path, TestSetup.MakeAdvancePokemon(nature: 3));

            var loaded = new PnbFile(path);
            loaded.Load();

            TeamMember neutral = TestSetup.MakeAdvancePokemon(nature: 0);
            Assert.IsTrue(loaded.Team[0].CurrentStats[StatTypes.nbAttack] > neutral.CurrentStats[StatTypes.nbAttack]);
            Assert.IsTrue(loaded.Team[0].CurrentStats[StatTypes.nbSpAttack] < neutral.CurrentStats[StatTypes.nbSpAttack]);
        }

        [TestMethod]
        public void Pnb_RejectsUnknownHeader() {
            string path = PathTo("bad.pnb");
            File.WriteAllBytes(path, Encoding.ASCII.GetBytes(" NOTPNB and some more data"));

            Assert.ThrowsException<InvalidDataException>(() => new PnbFile(path).Load());
        }

        [TestMethod]
        public void Pbx_RoundTrips() { // -- Code.bas WriteBoxPKMN: "BOX2.1" then Hex(Marker) & PKMN2Str
            string path = PathTo("stored.pbx");
            TeamMember member = TestSetup.MakeAdvancePokemon();
            member.MarkerNum = 11;

            new PbxFile(path) { BoxPokemon = new List<TeamMember> { member, TestSetup.MakeAdvancePokemon(TestSetup.Mew) } }.Save();
            Assert.AreEqual(6 + 2 * 36, new FileInfo(path).Length);

            var loaded = new PbxFile(path);
            loaded.Load();

            Assert.AreEqual(2, loaded.BoxPokemon.Count);
            Assert.AreEqual(11, loaded.BoxPokemon[0].MarkerNum);
            Assert.AreEqual("Testmon", loaded.BoxPokemon[0].Nickname);
            Assert.AreEqual(TestSetup.Mew, loaded.BoxPokemon[1].Pokemon.No);
            AssertSameStats(member, loaded.BoxPokemon[0]);
        }
    }
}
