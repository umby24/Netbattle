using Microsoft.VisualStudio.TestTools.UnitTesting;
using Netbattle.Common;
using Netbattle.Core.Battle;
using Netbattle.Core.Common;
using Netbattle.Core.Pokemon;

namespace NetbattleTests {
    [TestClass]
    public class TeamMemberTests {
        private static void AssertAdvanceStats(TeamMember member) {
            PokemonSpecies species = PokemonDatabase.BasePokemonMap[member.Pokemon.No];
            var rules = new ADVRules();

            Assert.AreEqual(rules.GetHp(species.BaseStats[StatTypes.nbHP], member.DVs[StatTypes.nbHP], member.EVs[StatTypes.nbHP], member.Level),
                member.CurrentStats[StatTypes.nbHP], "HP");

            foreach (StatTypes stat in StatSet.All.Where(s => s != StatTypes.nbHP)) {
                Assert.AreEqual(rules.GetStat(species.BaseStats[stat], member.DVs[stat], member.EVs[stat], member.Level,
                    NbMethods.GetNatureMod(member.NatureNum, (int)stat)), member.CurrentStats[stat], stat.ToString());
            }
        }

        [TestMethod]
        public void Initialize_AdvanceFillsEveryStatIncludingHp() {
            AssertAdvanceStats(TestSetup.MakeAdvancePokemon());
        }

        [TestMethod]
        public void Initialize_KeepsValidItem() { // -- Code.bas FillInPokeData
            Assert.AreEqual(Items.nbLeftovers, TestSetup.MakeAdvancePokemon().Item);
        }

        [TestMethod]
        public void Initialize_ClearsUnknownItem() {
            TeamMember member = TestSetup.MakeAdvancePokemon();
            member.Item = (Items)125;
            member.Initialize(CompatModes.nbFullAdvance);

            Assert.AreEqual(Items.nbNoItem, member.Item);
        }

        [TestMethod]
        public void Initialize_QuirkyDoesNotThrow() {
            AssertAdvanceStats(TestSetup.MakeAdvancePokemon(nature: 24));
        }

        [TestMethod]
        public void Initialize_GscKeepsDvsAtOrBelowFifteen() { // -- FillInPokeData caps, it doesn't overwrite.
            TeamMember member = TestSetup.MakeAdvancePokemon();
            member.DVs[StatTypes.nbAttack] = 12;
            member.DVs[StatTypes.nbDefense] = 7;
            member.DVs[StatTypes.nbSpeed] = 20;
            member.DVs[StatTypes.nbSpAttack] = 9;
            member.Initialize(CompatModes.nbTrueGSC);

            Assert.AreEqual(12, member.DVs[StatTypes.nbAttack]);
            Assert.AreEqual(7, member.DVs[StatTypes.nbDefense]);
            Assert.AreEqual(15, member.DVs[StatTypes.nbSpeed]);
            Assert.AreEqual(9, member.DVs[StatTypes.nbSpAttack]);
            Assert.AreEqual(0, member.DVs[StatTypes.nbSpDefense]);
            Assert.AreEqual(4 + 2 + 1, member.DVs[StatTypes.nbHP]); // -- Odd Def, Spd, SpAtk DVs
        }

        [TestMethod]
        public void Initialize_GscShinyChecksAttackDv() {
            TeamMember member = TestSetup.MakeAdvancePokemon();
            member.DVs[StatTypes.nbAttack] = 15;
            member.DVs[StatTypes.nbDefense] = 10;
            member.DVs[StatTypes.nbSpeed] = 10;
            member.DVs[StatTypes.nbSpAttack] = 10;
            member.Initialize(CompatModes.nbTrueGSC);
            Assert.IsTrue(member.Shiny);

            member.DVs[StatTypes.nbAttack] = 13;
            member.Initialize(CompatModes.nbTrueGSC);
            Assert.IsFalse(member.Shiny);
        }

        [TestMethod]
        public void Initialize_RbyUsesBaseSpecialForBothSpecials() {
            TeamMember member = TestSetup.MakeAdvancePokemon();
            member.DVs[StatTypes.nbSpAttack] = 15;
            member.Initialize(CompatModes.nbTrueRBY);

            int expected = new RBYRules().GetStat(member.Pokemon.BaseSpecial, 15, 0, 100, 0);
            Assert.AreEqual(expected, member.CurrentStats[StatTypes.nbSpAttack]);
            Assert.AreEqual(expected, member.CurrentStats[StatTypes.nbSpDefense]);
            Assert.AreEqual(Items.nbNoItem, member.Item);
        }

        [TestMethod]
        public void ApplyDefaultStats_MatchesTeamBuilderDefaults() {
            var advance = new TeamMember(PokemonDatabase.BasePokemonMap[TestSetup.Pikachu]);
            advance.ApplyDefaultStats(CompatModes.nbFullAdvance);
            advance.Initialize(CompatModes.nbFullAdvance);
            Assert.AreEqual(31, advance.DVs[StatTypes.nbSpeed]);
            Assert.AreEqual(85, advance.EVs[StatTypes.nbSpeed]);

            var gsc = new TeamMember(PokemonDatabase.BasePokemonMap[TestSetup.Pikachu]);
            gsc.ApplyDefaultStats(CompatModes.nbTrueGSC);
            gsc.Initialize(CompatModes.nbTrueGSC);
            Assert.AreEqual(15, gsc.DVs[StatTypes.nbSpeed]);
            Assert.AreEqual(0, gsc.EVs[StatTypes.nbSpeed]);
            Assert.AreEqual(15, gsc.DVs[StatTypes.nbHP]);
        }

        [TestMethod]
        public void Initialize_DoesNotChangeSharedSpecies() {
            PokemonSpecies species = PokemonDatabase.BasePokemonMap[TestSetup.Pikachu];
            int before = species.MaxHP;
            TestSetup.MakeAdvancePokemon();

            Assert.AreEqual(before, species.MaxHP);
        }

        [TestMethod]
        public void BinaryFormat_RoundTripsEveryField() {
            TeamMember original = TestSetup.MakeAdvancePokemon();
            original.Shiny = true;
            original.InBox = 5;
            byte[] packed = original.ToStringBytes();

            Assert.AreEqual(35, packed.Length);

            string nickname = System.Text.Encoding.ASCII.GetString(packed, 0, 15);
            TeamMember copy = TeamMember.FromBinary(NbMethods.BytesToBinary(packed.Skip(15).ToArray()), nickname);

            Assert.AreEqual("Testmon", copy.Nickname);
            Assert.AreEqual(original.Pokemon.No, copy.Pokemon.No);
            Assert.AreEqual(original.GameVersion, copy.GameVersion);
            Assert.AreEqual(original.Level, copy.Level);
            Assert.AreEqual(original.Item, copy.Item);
            Assert.AreEqual(original.NatureNum, copy.NatureNum);
            Assert.AreEqual(original.Shiny, copy.Shiny);
            Assert.AreEqual(original.InBox, copy.InBox);
            CollectionAssert.AreEqual(original.Move, copy.Move);

            foreach (StatTypes stat in StatSet.All) {
                Assert.AreEqual(original.DVs[stat], copy.DVs[stat], $"DV {stat}");
                Assert.AreEqual(original.EVs[stat], copy.EVs[stat], $"EV {stat}");
            }
        }

        [TestMethod]
        public void BinaryFormat_MatchesPkmn2StrBitLayout() {
            TeamMember member = TestSetup.MakeAdvancePokemon();
            string bits = NbMethods.BytesToBinary(member.ToStringBytes().Skip(15).ToArray());

            Assert.AreEqual(TestSetup.Pikachu, NbMethods.Bin2Dec(bits.Substring(0, 9)));
            Assert.AreEqual(100, NbMethods.Bin2Dec(bits.Substring(12, 7)));
            Assert.AreEqual(24, NbMethods.Bin2Dec(bits.Substring(149, 8)), "Last EV must keep its position in the final byte.");
            Assert.AreEqual("000", bits.Substring(157, 3), "Padding belongs after the data.");
        }

        [TestMethod]
        public void BinaryFormat_UnsetMoveIsWrittenAsZero() {
            TeamMember member = TestSetup.MakeAdvancePokemon();
            member.Move[3] = -1;
            string bits = NbMethods.BytesToBinary(member.ToStringBytes().Skip(15).ToArray());

            Assert.AreEqual(160, bits.Length);
            Assert.AreEqual(0, NbMethods.Bin2Dec(bits.Substring(70, 9)));
        }

        [TestMethod]
        public void FromBinary_EmptySlotIsNull() { // -- Str2PKMN: If X > UBound(BasePKMN) Or X = 0 Then Exit Function
            Assert.IsNull(TeamMember.FromBinary(new string('0', 160), ""));
        }

        [TestMethod]
        public void GetCurrentMoveset_LabelsMovesPerSpecies() {
            // -- The same Move is learned by different means on different species, so a shared object can't carry the source.
            var sources = PokemonDatabase.BasePokemon
                .SelectMany(p => p.Moves.Sets)
                .SelectMany(set => set.Value.Select(move => (set.Key, move)));

            foreach ((Netbattle.Core.Database.MoveTypes source, Move move) in sources)
                Assert.AreEqual(source, move.Source, move.Name);
        }

        [TestMethod]
        public void MoveEquality_ComparesById() {
            var database = PokemonDatabase.BasePokemonMap[TestSetup.Pikachu].Moves.Sets[Netbattle.Core.Database.MoveTypes.nbAdvLevel][0];
            Assert.IsTrue(Netbattle.Core.Database.MoveDatabase.Moves[database.ID].Equals(database));
        }
    }
}
