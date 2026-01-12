using System.Collections.Generic;
using Netbattle.Common;

namespace Netbattle.Database {
    /// <summary>
    /// Loader and manager for the type effectiveness database
    /// </summary>
    public class TypeDatabase {
        public static float[,] BattleMatrix = new float[18, 18]; //Type effectiveness chart - (AttackType,DefendType)
        public static string[] AttributeText = new string[78];
        
        public static void Load() {
            var typeDb = new CdbFile("TypeDB.cdb");
            typeDb.Load();

            Logger.Log(LogType.Verbose, "Type database read and decompressed.");

            foreach (string[] entry in typeDb.LineContent) {
                ParseFileEntry(entry);
            }
            
	        PopulateAttributeText();
            Logger.Log(LogType.Verbose, "Type database loaded successfully.");
        }

        private static void PopulateAttributeText() {
			AttributeText[0] = "No Trait";
			AttributeText[1] = "Stench";
			AttributeText[2] = "Drizzle";
			AttributeText[3] = "Speed Boost";
			AttributeText[4] = "Battle Armor";
			AttributeText[5] = "Sturdy";
			AttributeText[6] = "Damp";
			AttributeText[7] = "Limber";
			AttributeText[8] = "Sand Veil";
			AttributeText[9] = "Static";
			AttributeText[10] = "Volt Absorb";
			AttributeText[11] = "Water Absorb";
			AttributeText[12] = "Oblivious";
			AttributeText[13] = "Cloud Nine";
			AttributeText[14] = "Compoundeyes";
			AttributeText[15] = "Insomnia";
			AttributeText[16] = "Color Change";
			AttributeText[17] = "Immunity";
			AttributeText[18] = "Flash Fire";
			AttributeText[19] = "Shield Dust";
			AttributeText[20] = "Own Tempo";
			AttributeText[21] = "Suction Cups";
			AttributeText[22] = "Intimidate";
			AttributeText[23] = "Shadow Tag";
			AttributeText[24] = "Rough Skin";
			AttributeText[25] = "Wonder Guard";
			AttributeText[26] = "Levitate";
			AttributeText[27] = "Effect Spore";
			AttributeText[28] = "Synchronize";
			AttributeText[29] = "Clear Body";
			AttributeText[30] = "Natural Cure";
			AttributeText[31] = "Lightning Rod";
			AttributeText[32] = "Serene Grace";
			AttributeText[33] = "Swift Swim";
			AttributeText[34] = "Chlorophyll";
			AttributeText[35] = "Illuminate";
			AttributeText[36] = "Trace";
			AttributeText[37] = "Huge Power";
			AttributeText[38] = "Poison Point";
			AttributeText[39] = "Inner Focus";
			AttributeText[40] = "Magma Armor";
			AttributeText[41] = "Water Veil";
			AttributeText[42] = "Magnet Pull";
			AttributeText[43] = "Soundproof";
			AttributeText[44] = "Rain Dish";
			AttributeText[45] = "Sand Stream";
			AttributeText[46] = "Pressure";
			AttributeText[47] = "Thick Fat";
			AttributeText[48] = "Early Bird";
			AttributeText[49] = "Flame Body";
			AttributeText[50] = "Run Away";
			AttributeText[51] = "Keen Eye";
			AttributeText[52] = "Hyper Cutter";
			AttributeText[53] = "Pickup";
			AttributeText[54] = "Truant";
			AttributeText[55] = "Hustle";
			AttributeText[56] = "Cute Charm";
			AttributeText[57] = "Plus";
			AttributeText[58] = "Minus";
			AttributeText[59] = "Forecast";
			AttributeText[60] = "Sticky Hold";
			AttributeText[61] = "Shed Skin";
			AttributeText[62] = "Guts";
			AttributeText[63] = "Marvel Scale";
			AttributeText[64] = "Liquid Ooze";
			AttributeText[65] = "Overgrow";
			AttributeText[66] = "Blaze";
			AttributeText[67] = "Torrent";
			AttributeText[68] = "Swarm";
			AttributeText[69] = "Rock Head";
			AttributeText[70] = "Drought";
			AttributeText[71] = "Arena Trap";
			AttributeText[72] = "Vital Spirit";
			AttributeText[73] = "White Smoke";
			AttributeText[74] = "Pure Power";
			AttributeText[75] = "Shell Armor";
			AttributeText[76] = "Cacophony";
			AttributeText[77] = "Air Lock";
        }
        private static void ParseFileEntry(IReadOnlyList<string> entry) {
            for (var i = 0; i < 17; i++) { // -- OG: X = 1 -> 17.
                BattleMatrix[int.Parse(entry[0]), i] = float.Parse(entry[i + 1]);
            }
        }
    }
}
