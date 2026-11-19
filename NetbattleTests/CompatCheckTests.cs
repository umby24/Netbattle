using Microsoft.VisualStudio.TestTools.UnitTesting;
using Netbattle.Common;
using Netbattle.Core.Common;
using Netbattle.Core.Database;

namespace NetbattleTests {
    [TestClass]
    public class CompatCheckTests {
        private const int Treecko = 252;
        private const int RbyIndex = 0, TrueRbyIndex = 5;

        private static TeamMember Plain(int number) => new TeamMember(PokemonDatabase.BasePokemonMap[number]) {
            Item = Items.nbNoItem,
            Move = new[] { 0, 0, 0, 0 }
        };

        [TestMethod]
        public void EmptyTeamIsCompatibleWithEverything() { // -- CompatCheck starts every TempCheck at True.
            bool[] modes = BattleSystem.CompatCheck(new TeamMember?[6]);
            Assert.IsTrue(modes.All(m => m));
        }

        [TestMethod]
        public void EmptySlotsAndMovesAreSkipped() { // -- VB: CheckPKMN(X).No > 0, Move(Y) > 0
            var team = new TeamMember?[] { Plain(TestSetup.Pikachu), null, null, null, null, null };
            bool[] modes = BattleSystem.CompatCheck(team);
            Assert.IsTrue(modes[RbyIndex]);
        }

        [TestMethod]
        public void LaterMembersCannotReenableAMode() { // -- VB only ever sets TempCheck(X) = False.
            var adv = new TeamMember?[] { Plain(Treecko), Plain(TestSetup.Pikachu) };
            var advLast = new TeamMember?[] { Plain(TestSetup.Pikachu), Plain(Treecko) };

            foreach (TeamMember?[] team in new[] { adv, advLast }) {
                bool[] modes = BattleSystem.CompatCheck(team);
                Assert.IsFalse(modes[RbyIndex], "Treecko isn't in RBY.");
                Assert.IsFalse(modes[TrueRbyIndex], "Treecko isn't in RBY.");
            }
        }
    }
}
