using Microsoft.VisualStudio.TestTools.UnitTesting;
using Netbattle.Common;
using Netbattle.Core.Common;
using Netbattle.Core.Database;

[assembly: DoNotParallelize] // -- The databases and UserSettings are global static state.

namespace NetbattleTests {
    [TestClass]
    public static class TestSetup {
        public const int Pikachu = 25;
        public const int Mew = 151;

        [AssemblyInitialize]
        public static void LoadDatabases(TestContext context) {
            Directory.SetCurrentDirectory(AppContext.BaseDirectory); // -- Databases load by relative path.
            new Logger().Setup();
            MoveDatabase.Load();
            TypeDatabase.Load();
            PokemonDatabase.Load();
            BattleSystem.GenerateStatTable();
        }

        /// <summary>
        /// A level 100 Pokemon with distinct DVs and EVs, so that swapped fields show up as wrong values.
        /// </summary>
        public static TeamMember MakeAdvancePokemon(int number = Pikachu, byte nature = 3) {
            var member = new TeamMember(PokemonDatabase.BasePokemonMap[number]) {
                Nickname = "Testmon",
                Level = 100,
                NatureNum = nature,
                Item = Items.nbLeftovers,
                Move = new[] { 85, 57, 94, 63 }
            };

            member.DVs[StatTypes.nbHP] = 31;
            member.DVs[StatTypes.nbAttack] = 30;
            member.DVs[StatTypes.nbDefense] = 29;
            member.DVs[StatTypes.nbSpeed] = 28;
            member.DVs[StatTypes.nbSpAttack] = 27;
            member.DVs[StatTypes.nbSpDefense] = 26;
            member.EVs[StatTypes.nbHP] = 4;
            member.EVs[StatTypes.nbAttack] = 252;
            member.EVs[StatTypes.nbDefense] = 8;
            member.EVs[StatTypes.nbSpeed] = 200;
            member.EVs[StatTypes.nbSpAttack] = 16;
            member.EVs[StatTypes.nbSpDefense] = 24;
            member.Initialize(CompatModes.nbFullAdvance);
            return member;
        }
    }
}
