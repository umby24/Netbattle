using Microsoft.VisualStudio.TestTools.UnitTesting;
using Netbattle.Core.Battle;
using Netbattle.Core.Common;
using Netbattle.Core.Database;
using Netbattle.Core.Pokemon;

namespace NetbattleTests {
    [TestClass]
    public class RulesetTests {
        // -- Expected values are worked by hand from Code.bas GetAdvStat / GetAdvHP, which floor rather than round.
        [DataTestMethod]
        [DataRow(100, 31, 0, 50, 0, 120)]
        [DataRow(50, 0, 0, 100, 0, 105)]
        [DataRow(50, 0, 0, 100, 1, 115)] // -- 105 * 1.1 = 115.5, a round would give 116.
        [DataRow(50, 0, 0, 100, -1, 94)] // -- 105 * 0.9 = 94.5
        [DataRow(90, 31, 252, 100, 1, 306)] // -- 279 * 1.1 = 306.9
        [DataRow(100, 31, 0, 50, -1, 107)] // -- 120 * 0.9 lands just under 108: VB's p is a Single.
        public void AdvGetStat_FloorsLikeVb(int baseVal, int iv, int ev, int level, int pmod, int expected) {
            Assert.AreEqual(expected, new ADVRules().GetStat(baseVal, iv, ev, level, pmod));
        }

        [DataTestMethod]
        [DataRow(100, 31, 0, 50, 175)] // -- 115.5 floors, a round would give 176.
        [DataRow(100, 31, 252, 100, 404)]
        public void AdvGetHp_FloorsLikeVb(int baseVal, int iv, int ev, int level, int expected) {
            Assert.AreEqual(expected, new ADVRules().GetHp(baseVal, iv, ev, level));
        }

        [TestMethod]
        public void AdvGetHp_ShedinjaIsAlwaysOne() {
            Assert.AreEqual(1, new ADVRules().GetHp(1, 31, 252, 100));
        }

        [DataTestMethod]
        [DataRow(100, 15, 100, 298)]
        [DataRow(100, 0, 50, 136)]
        public void GscGetStat_MatchesVb(int baseVal, int dv, int level, int expected) {
            Assert.AreEqual(expected, new GSCRules().GetStat(baseVal, dv, 0, level, 0));
        }

        [TestMethod]
        public void GscGetHp_MatchesVb() {
            Assert.AreEqual((int)(100 * (100 + 15 + 31.9 + 50) / 50) + 10, new GSCRules().GetHp(100, 15, 0, 100));
        }

        [TestMethod]
        public void AdvHiddenPowerStrength_UsesSecondBitOfEachDv() { // -- Code.bas HiddenPowerStrengthAdv
            var allThirtyOne = new StatSet();
            foreach (StatTypes stat in StatSet.All)
                allThirtyOne[stat] = 31;

            Assert.AreEqual(70, new ADVRules().HiddenPowerStrength(allThirtyOne));
            Assert.AreEqual(30, new ADVRules().HiddenPowerStrength(new StatSet()));
        }

        [TestMethod]
        public void AdvHiddenPowerType_AllOddDvsIsDark() {
            var dvs = new StatSet();
            foreach (StatTypes stat in StatSet.All)
                dvs[stat] = 31;

            Assert.AreEqual(Elements.nbDark, new ADVRules().HiddenPowerType(dvs));
        }

        [TestMethod]
        public void GscHiddenPower_MaxDvs() { // -- Code.bas HiddenPowerStrength / HiddenPowerType
            var dvs = new StatSet();
            foreach (StatTypes stat in StatSet.All)
                dvs[stat] = 15;

            Assert.AreEqual(70, new GSCRules().HiddenPowerStrength(dvs));
            Assert.AreEqual(Elements.nbDark, new GSCRules().HiddenPowerType(dvs));
        }

        [DataTestMethod]
        [DataRow(CompatModes.nbRBYTrade, typeof(RBYRules))]
        [DataRow(CompatModes.nbTrueRBY, typeof(RBYRules))]
        [DataRow(CompatModes.nbGSCTrade, typeof(GSCRules))]
        [DataRow(CompatModes.nbTrueGSC, typeof(GSCRules))]
        [DataRow(CompatModes.nbTrueRuSa, typeof(ADVRules))]
        [DataRow(CompatModes.nbFullAdvance, typeof(ADVRules))]
        [DataRow(CompatModes.nbModAdv, typeof(ADVRules))]
        public void GetGenerationRuleset_MapsEveryMode(CompatModes mode, Type expected) {
            Assert.IsInstanceOfType(BattleSystem.GetGenerationRuleset(mode), expected);
        }

        [TestMethod]
        public void NatureTable_CoversAllTwentyFiveNatures() { // -- PokeLoader.frm: For X = 0 To 24
            Assert.AreEqual(25, BattleSystem.NatureStats.Length);

            for (var i = 0; i < 25; i++) {
                Assert.IsNotNull(BattleSystem.NatureStats[i], $"Nature {i}");
                Assert.AreEqual(0, BattleSystem.NatureStats[i].StatChg.Sum(), $"Nature {i}");
            }
        }

        [TestMethod]
        public void NatureTable_Adamant() {
            int[] change = BattleSystem.NatureStats[3].StatChg;

            Assert.AreEqual(1, change[(int)StatTypes.nbAttack]);
            Assert.AreEqual(-1, change[(int)StatTypes.nbSpAttack]);
        }
    }
}
