using System.Collections.Generic;
using Netbattle.Common;

namespace Netbattle.Database {
    /// <summary>
    /// Loader and manager for the type effectiveness database
    /// </summary>
    public class TypeDatabase {
        public static float[,] BattleMatrix = new float[18, 18]; //Type effectiveness chart - (AttackType,DefendType)

        public static void Load() {
            var typeDb = new CdbFile("TypeDB.cdb");
            typeDb.Load();

            Logger.Log(LogType.Verbose, "Type database read and decompressed.");

            foreach (string[] entry in typeDb.LineContent) {
                ParseFileEntry(entry);
            }

            Logger.Log(LogType.Verbose, "Type database loaded successfully.");
        }

        private static void ParseFileEntry(IReadOnlyList<string> entry) {
            for (var i = 0; i < 17; i++) { // -- OG: X = 1 -> 17.
                BattleMatrix[int.Parse(entry[0]), i] = float.Parse(entry[i + 1]);
            }
        }
    }
}
