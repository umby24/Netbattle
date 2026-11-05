using Microsoft.VisualStudio.TestTools.UnitTesting;
using Netbattle.Common;
using Netbattle.Core.Battle;
using Netbattle.Core.Common;
using Netbattle.Core.Database;
using Netbattle.Core.Pokemon;

namespace NetbattleTests {
    [TestClass]
    public class NbMethodsTests {
        [DataTestMethod]
        [DataRow(1, 9, "000000001")]
        [DataRow(2, 9, "000000010")]
        [DataRow(30, 9, "000011110")]
        [DataRow(100, 9, "001100100")]
        [DataRow(0, 5, "00000")]
        public void Dec2Bin_PadsToLength(int value, int length, string expected) {
            Assert.AreEqual(expected, NbMethods.Dec2Bin(value, length));
        }

        [TestMethod]
        public void Dec2Bin_KeepsRightmostBitsWhenTooWide() { // -- Right$(sByte(X), Fixed)
            Assert.AreEqual("11", NbMethods.Dec2Bin(255, 2));
        }

        [DataTestMethod]
        [DataRow("000000001", 1)]
        [DataRow("001100100", 100)]
        [DataRow("", 0)]
        public void Bin2Dec_ParsesBitString(string bits, int expected) {
            Assert.AreEqual(expected, NbMethods.Bin2Dec(bits));
        }

        [TestMethod]
        public void BinaryToBytes_PadsPartialByteOnTheRight() { // -- Code.bas Bin2Chr
            CollectionAssert.AreEqual(new byte[] { 0xFF, 0xA8 }, NbMethods.BinaryToBytes("1111111110101"));
        }

        [TestMethod]
        public void BinaryConversions_RoundTrip() {
            var original = new byte[] { 0x19, 0xBC, 0x85, 0x07, 0xFF };
            CollectionAssert.AreEqual(original, NbMethods.BinaryToBytes(NbMethods.BytesToBinary(original)));
        }

        [TestMethod]
        public void Sid_RoundTrips() {
            // -- The first SID character comes from every 20th bit, so only a decompressed SID is valid input.
            string sid = NbMethods.DecompressSid(new byte[] { 0x19, 0xBC, 0x85, 0x07, 0x3A, 0xF1, 0x81, 0xFF, 0xFB, 0xDE, 0xF0, 0x27, 0xE0 });
            byte[] compressed = NbMethods.CompressSid(sid);

            Assert.AreEqual(13, compressed.Length);
            Assert.AreEqual(sid, NbMethods.DecompressSid(compressed));
        }

        [DataTestMethod]
        [DataRow(255, 2, "FF")]
        [DataRow(10, 1, "A")]
        [DataRow(10, 2, "0A")]
        [DataRow(0x1FF, 2, "FF")]
        public void FixedHex_MatchesVbHex(int value, int length, string expected) {
            Assert.AreEqual(expected, NbMethods.FixedHex(value, length));
        }

        [TestMethod]
        public void GetNatureMod_MatchesNatureTable() {
            Assert.AreEqual(1, NbMethods.GetNatureMod(3, 1)); // -- Adamant, +Atk
            Assert.AreEqual(-1, NbMethods.GetNatureMod(3, 4)); // -- -SpAtk
            Assert.AreEqual(0, NbMethods.GetNatureMod(24, 5)); // -- Quirky is neutral
            Assert.AreEqual(0, NbMethods.GetNatureMod(255, 1));
        }

        [TestMethod]
        public void GetPokeRank_AppliesNature() {
            int adamant = NbMethods.GetPokeRank(TestSetup.MakeAdvancePokemon(nature: 3));
            int hardy = NbMethods.GetPokeRank(TestSetup.MakeAdvancePokemon(nature: 0));

            Assert.AreNotEqual(hardy, adamant);
        }

        [TestMethod]
        public void GetPokeRank_SkewsLegendaries() { // -- Code.bas GetPokeRank: Level * 2.5 tier
            TeamMember mew = TestSetup.MakeAdvancePokemon(TestSetup.Mew);
            var rules = new ADVRules();

            double stats = rules.GetHp(100, mew.DVs[StatTypes.nbHP], mew.EVs[StatTypes.nbHP], 100) / 1.5;
            foreach (StatTypes stat in StatSet.All.Where(s => s != StatTypes.nbHP))
                stats += rules.GetStat(100, mew.DVs[stat], mew.EVs[stat], 100, NbMethods.GetNatureMod(mew.NatureNum, (int)stat));

            var matrixAdjust = 0;
            for (var attacker = 1; attacker <= 17; attacker++) {
                float damage = TypeDatabase.BattleMatrix[attacker, (int)Elements.nbPsychc];
                matrixAdjust += damage == 2 ? -35 : damage == 0.5f ? 25 : damage == 0 ? 75 : 0;
            }

            Assert.AreEqual((int)Math.Round(stats + matrixAdjust + 250), NbMethods.GetPokeRank(mew));
        }

        [TestMethod]
        public void GetTeamRank_IsCappedAtOneHundred() {
            var team = new TeamMember[6];
            for (var i = 0; i < 6; i++)
                team[i] = TestSetup.MakeAdvancePokemon();

            int rank = NbMethods.GetTeamRank(team);
            Assert.IsTrue(rank is > 0 and < 100, $"Rank was {rank}");
        }

        [TestMethod]
        public void GetTeamRank_EmptyTeamIsZero() {
            Assert.AreEqual(0, NbMethods.GetTeamRank(new TeamMember[6]));
        }
    }
}
