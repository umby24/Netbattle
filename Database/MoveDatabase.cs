using System.Collections.Generic;
using Netbattle.Common;

namespace Netbattle.Database {
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
