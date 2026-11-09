using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Netbattle.Core.Common;
using Netbattle.Network;

namespace NetbattleTests {
    [TestClass]
    public class ChallengeTests {
        [TestMethod]
        public void MakeBinArray_RuleNIsBitNMinusOne() { // -- Bit(LBound) = 1, doubling upwards
            var rules = new bool[ChallengeRules.RuleCount + 1];
            rules[(int)RuleEnum.nbSleep] = true;     // -- 1
            rules[(int)RuleEnum.nbRandbat] = true;   // -- 32
            rules[(int)RuleEnum.nbPresentRule] = true; // -- 2048

            Assert.AreEqual(1 + 32 + 2048, ChallengeRules.MakeBinArray(rules));
        }

        [TestMethod]
        public void ReadBinArray_UndoesMakeBinArray() {
            bool[] rules = ChallengeRules.ReadBinArray(0x821);

            Assert.IsTrue(rules[(int)RuleEnum.nbSleep]);
            Assert.IsTrue(rules[(int)RuleEnum.nbRandbat]);
            Assert.IsTrue(rules[(int)RuleEnum.nbPresentRule]);
            Assert.AreEqual(3, rules.Count(r => r));
            Assert.AreEqual(0x821, ChallengeRules.MakeBinArray(rules));
        }

        [TestMethod]
        public void FixedHex_PadsAndKeepsTheRightmostDigits() {
            Assert.AreEqual("00A", ChallengeRules.FixedHex(10, 3));
            Assert.AreEqual("23", ChallengeRules.FixedHex(0x123, 2)); // -- Right(Temp, Length)
        }

        [TestMethod]
        public void CompatVersion_MatchesVb() {
            Assert.AreEqual(BattleModes.nbRBYBattle, ChallengeRules.CompatVersion(5));
            Assert.AreEqual(BattleModes.nbGSCBattle, ChallengeRules.CompatVersion(6));
            Assert.AreEqual(BattleModes.nbAdvBattle, ChallengeRules.CompatVersion(4));
        }

        [TestMethod]
        public void SentChallenge_MatchesVbFormat() {
            // -- "CHLN:" & FixedHex(ChallengeNumber, 3) & GameVersion & FixedHex(Mode, 1) & FixedHex(Terrain, 1) & FixedHex(Rules, 8)
            var packet = new ChallengeReceived {
                PlayerId = 26,
                GameVersion = 4,
                Settings = new ChallengeSettings { Mode = 3, Terrain = 10, Rules = ChallengeRules.ReadBinArray(0x21) }
            };
            var buffer = new ByteBuffer();
            packet.Write(buffer);

            Assert.AreEqual("CHLN:01A43A00000021", Encoding.ASCII.GetString(buffer.GetAllBytes()));
        }

        [TestMethod]
        public void ReceivedChallenge_ReadsSlotModeTerrainRulesAndSpeed() {
            var buffer = new ByteBuffer();
            buffer.WriteString("01A2B000008200150");
            var packet = new ChallengeReceived();
            packet.Read(buffer);

            Assert.AreEqual(26, packet.PlayerId);
            Assert.AreEqual(2, packet.Settings.Mode);
            Assert.AreEqual(11, packet.Settings.Terrain);
            Assert.IsTrue(packet.Settings.Rules[(int)RuleEnum.nbRandbat]);
            Assert.IsTrue(packet.Settings.Rules[(int)RuleEnum.nbPresentRule]);
            Assert.AreEqual("0150", packet.Speed);
        }

        [TestMethod]
        public void AcceptedChallenge_MatchesVbFormat() {
            // -- "PACC:" & FixedHex(ChallengeNumber, 3) & FixedHex(GameMode, 1) & FixedHex(ChallTerrain, 1) & FixedHex(Rules, 8)
            var packet = new PlayerAcceptedChallenge {
                PlayerId = 300,
                Settings = new ChallengeSettings { Mode = 2, Terrain = 0, Rules = ChallengeRules.ReadBinArray(0x800) }
            };
            var buffer = new ByteBuffer();
            packet.Write(buffer);

            Assert.AreEqual("PACC:12C2000000800", Encoding.ASCII.GetString(buffer.GetAllBytes()));
        }
    }
}
