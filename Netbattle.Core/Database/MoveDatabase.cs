using Netbattle.Common;
using Netbattle.Core.Common;
using Netbattle.Core.Files;

namespace Netbattle.Core.Database {
    public enum MoveTypes
    {
        nbRBYLevel = 1,
        nbRBYTM = 2,
        nbGSCLevel = 3,
        nbGSCTM = 4,
        nbGSCEgg = 5,
        nbGSCTutor = 6,
        nbGSCSpecial = 7,
        nbAdvLevel = 8,
        nbAdvTM = 9,
        nbAdvEgg = 10,
        nbAdvTutor = 11,
        nbAdvSpecial = 12,
        nbAdvFL = 13
    }
    
    /// <summary>
    /// Loader and manager of the Pokemon Move Database.
    /// </summary>
    public class MoveDatabase {
        public static Dictionary<int, Move> Moves = new Dictionary<int, Move>();

        public static void Load() {
            var moveDb = new CdbFile("MoveDB.cdb");
            moveDb.Load();
            Logger.Log(LogType.Verbose, "MoveDB Read and decompressed.");

            foreach (string[] entry in moveDb.LineContent) {
                ParseFileEntry(entry);
            }

            Logger.Log(LogType.Verbose, $"MoveDB load complete. {Moves.Count} total moves.");
        }
        
        /// <summary>
        /// Determines the battle mode based on the compatibility version
        /// </summary>
        private BattleModes CompatVersion(CompatModes gameVersion)
        {
            switch (gameVersion)
            {
                case CompatModes.nbTrueRBY:
                case CompatModes.nbTrueGSC:
                case CompatModes.nbRBYTrade:
                case CompatModes.nbGSCTrade:
                    return BattleModes.nbGSCBattle;

                case CompatModes.nbTrueRuSa:
                case CompatModes.nbFullAdvance:
                case CompatModes.nbModAdv:
                    return BattleModes.nbAdvBattle;

                default:
                    return BattleModes.nbGSCBattle;
            }
        }

        /// <summary>
        /// Checks if a moveset violates database-defined illegal combinations
        /// </summary>
        private bool DBIllegalCheck(int[] moves, string illegals)
        {
            if (string.IsNullOrEmpty(illegals))
            {
                return true;
            }

            var legality = true;
            var sets = illegals.Split('|');
            foreach (var set in sets) {
                var tempSet = set.Split('+');
                legality = HasMoves(moves, tempSet.Select(short.Parse).ToList());
                if (legality)
                    break;
            }
            
            return !legality; // If any illegal set is found, return false
        }
        
        private bool HasMoves(int[] moves, List<short> checkMoves) {
            int instances = checkMoves.Sum(move => moves.Count(m => m == move));
            return instances == checkMoves.Count;
        }
        
        /// <summary>
        /// Checks if egg move combinations are legal for breeding
        /// </summary>
        private bool BreedCheck(int pokemonNo, int[] moves, CompatModes gameVersion)
        {
            // Implementation depends on breeding chain data structure
            var basePkmn = PokemonDatabase.BasePokemon[pokemonNo];
            string temp;
            if (CompatVersion(gameVersion) == BattleModes.nbGSCBattle) {
                temp = basePkmn.BreedIllegals[0];
                if (gameVersion == CompatModes.nbTrueGSC) temp += basePkmn.BreedIllegals[1];
            }
            else {
                temp = basePkmn.BreedIllegals[2];
                if (gameVersion == CompatModes.nbTrueGSC) temp += basePkmn.BreedIllegals[3];
            }
            
            return DBIllegalCheck(moves, temp);
        }
        
        private static void ParseFileEntry(IReadOnlyList<string> entry) {
            var result = new Move {
                ID = short.Parse(entry[0]),
                Name = entry[1],
                Type = (Elements)int.Parse(entry[2]),
                Power = short.Parse(entry[3]),
                Accuracy = byte.Parse(entry[4]),
                PP = byte.Parse(entry[5]),
                SpecialPercent = byte.Parse(entry[6]),
                SpecialEffect = byte.Parse(entry[7]),
                Target = (MoveTargets)int.Parse(entry[8]),
                Text = entry[9],
                WorksRight = int.Parse(entry[10]) > 0,
                BrightPowder = int.Parse(entry[11]) > 0,
                KingsRock = int.Parse(entry[12]) > 0,
                RBYMove = int.Parse(entry[13]) > 0,
                GSCMove = int.Parse(entry[14]) > 0,
                AdvMove = int.Parse(entry[15]) > 0,
                HitsTeam = int.Parse(entry[16]) > 0,
                SelfMove = int.Parse(entry[17]) > 0,
                OldTM = entry[18],
                NewTM = entry[19],
                ADVTM = entry[20],
                SubstituteBlocks = int.Parse(entry[21]) > 0,
                HitsAll = int.Parse(entry[22]) > 0,
                SoundMove = int.Parse(entry[23]) > 0,
                PhysMove = int.Parse(entry[24]) > 0,
                MagicCoat = int.Parse(entry[25]) > 0
            };

            Moves.Add(result.ID, result);
        }
    }
}
