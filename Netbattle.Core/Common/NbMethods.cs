using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Netbattle.Common;
using Netbattle.Core.Database;

namespace Netbattle.Core.Common {
    /// <summary>
    /// Static class to hold sane-versions of Netbattle's obscure methods.
    /// </summary>
    public static class NbMethods {
        public static byte[] StringToByteArray(string hex) {
            return Enumerable.Range(0, hex.Length)
                .Where(x => x % 2 == 0)
                .Select(x => Convert.ToByte(hex.Substring(x, 2), 16))
                .ToArray();
        }

        public static int MakeBinArray(ref bool[] inArray) {
            var bit = new int[inArray.Length];
            int result = 0;
            
            bit[0] = 1;
            for (var i = 1; i < inArray.Length; i++) {
                bit[i] = bit[i - 1]*2;
            }

            for (var i = 0; i < inArray.Length; i++) {
                if (inArray[i]) result += bit[i];
            }

            return result;
        }
        
        /// <summary>
        /// Convert an integer into a bit string.
        /// </summary>
        /// <param name="input"></param>
        /// <param name="length"></param>
        public static string Dec2Bin(int input, int length) {
            string asBits = Convert.ToString(input, 2).PadLeft(length, '0');
            return asBits.Substring(asBits.Length - length);
        }

        public static int Bin2Dec(string input) {
            if (input.Length < 32)
                input = input.PadLeft(32, '0');

            var intBytes = BinaryToBytes(input);
            Array.Reverse(intBytes); // -- I shouldn't need this.. why do I need this? Why is bitconverter suddenly big endian??
            int result = BitConverter.ToInt32(intBytes, 0);
            return result;
        }

        /// <summary>
        /// This method Takes a bit-string and converts it to bytes.
        /// </summary>
        /// <param name="bitString"></param>
        public static byte[] Bin2Chr(string bitString) {
            int padding = 8 - (bitString.Length%8);

            if (padding == 8)
                padding = 0;

            // -- Pad with zeros if the length is not a multiple of 8.
            bitString = bitString.PadRight(bitString.Length + padding, '0');
            

            int arraySize = bitString.Length/8;
            var result = new byte[arraySize];
            var workingIndex = 0;

            for (var i = 0; i < arraySize; i++) {
                string temp = bitString.Substring(workingIndex, 8); // -- Grab 8 bits
                workingIndex += 8;
                result[i] = (byte)Convert.ToInt32(temp, 2); // -- Convert the 8 bits into a byte value.
            }

            return result;
        }

        /// <summary>
        /// Converts a number into a hex string, padding or removing bytes to match a specific length.
        /// </summary>
        /// <param name="number"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public static string FixedHex(int number, int length) {
            string hexed = Convert.ToString(number, 16).ToUpperInvariant(); // -- Convert the number to base-16 (HEX)

            if (hexed.Length == length)
                return hexed;

            if (hexed.Length <= length)
                return hexed.PadLeft(length, '0'); // -- length > hexed.length

            int diff = hexed.Length - length;
            hexed = hexed.Substring(diff, hexed.Length - diff);
            return hexed;
        }
        
        /// <summary>
        /// Takes a byte array and converts it to a binary string.
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns></returns>
        public static string BytesToBinary(byte[] bytes) {
            var build = "";

            foreach (byte b in bytes) {
                string chr = Convert.ToString(b, 2);
                chr = chr.PadLeft(8, '0');
                build += chr;
            }

            return build;
        }
        
        /// <summary>
        /// Takes a binary string and converts it to a byte array.
        /// </summary>
        /// <param name="binary">A string of only 0s and 1s.</param>
        /// <returns></returns>
        public static byte[] BinaryToBytes(string binary) {
            return Bin2Chr(binary);
        }

        /// <summary>
        /// A Recreation of Netbattle's 'Decompress SID' method. Takes 13 bytes, and produces the 21-byte system ID.
        /// </summary>
        /// <param name="sidBytes"></param>
        public static string DecompressSid(byte[] sidBytes) {
            string asBinary = BytesToBinary(sidBytes).Substring(0, 100); // -- Breaks the bytes down to their individual bits, in a full string.
            var result = "";
            var temp = "";
            for (var i = 1; i < 6; i++) { // -- Takes every 20th bit, and places it at the front of the build order..
                temp += asBinary.Substring((i*20) - 1, 1);
            }
            asBinary = temp + asBinary;
            for (var i = 1; i < 22; i++) { // -- Takes each 5 bit set, treating them as their own values.
                string sub = asBinary.Substring((i*5) - 5, 5);
                int val = Convert.ToInt32(sub, 2); // -- Converts them to a value..
                val += (val > 8 ? 56 : 49); // -- Adds either 56 or 49 to them.
                result += (char) val; // -- and that's the byte value of the character!
            }

            return result;
        }

        // -- Legal Charset: (1-9, A-W)
        // -- I.E. Reverse SID Creation :D only 15 years later...
        public static byte[] CompressSid(string output) {
            byte[] stringBytes = Encoding.ASCII.GetBytes(output);
            var builder = "";

            for (var i = 0; i < 21; i++) {
                int val;

                if (stringBytes[i] >= 65) // -- Gets the proper lower val..
                    val = stringBytes[i] - 56;
                else
                    val = stringBytes[i] - 49;

                // -- Convert it to bits..
                string valBits = Convert.ToString(val, 2);
                valBits = valBits.PadLeft(5, '0');
                builder += valBits;
            }

            // -- Now we need to take the first 5 bits, and disperse them amungst the other bits..
            var temp = new StringBuilder(builder.Substring(5, builder.Length - 5));
            //for (int i = 1; i < 6; i++) { // -- Takes every 20th bit, and places it at the front of the build order..
            //    temp[(i*20) - 1] = builder[i - 1];
            //}
            return BinaryToBytes(temp.ToString());
        }
        
        /// <summary>
        /// Generates a numerical power score for this pokemon.
        /// The score is essentially just jamming all the stats together into a big number.
        /// </summary>
        public static int GetPokeRank(TeamMember p)
        {
            if (p == null || p.Pokemon.No == 0)
                return 0;
            int total = 0;
            var basePkmn = PokemonDatabase.BasePokemonMap[p.Pokemon.No];

            var ruleSet = BattleSystem.GetGenerationRuleset(p.GameVersion);
            // -- RBY and GSC share one Special DV.
            StatTypes spDefDv = BattleSystem.IsAdvance(p.GameVersion) ? StatTypes.nbSpDefense : StatTypes.nbSpAttack;
            double hp = ruleSet.GetHp(basePkmn.BaseStats[StatTypes.nbHP], p.DVs[StatTypes.nbHP], p.EVs[StatTypes.nbHP],
                p.Level) / 1.5;
            var atk = ruleSet.GetStat(basePkmn.BaseStats[StatTypes.nbAttack], p.DVs[StatTypes.nbAttack],
                p.EVs[StatTypes.nbAttack], p.Level, GetNatureMod(p.NatureNum, 1));

            var def = ruleSet.GetStat(basePkmn.BaseStats[StatTypes.nbDefense], p.DVs[StatTypes.nbDefense],
                p.EVs[StatTypes.nbDefense], p.Level, GetNatureMod(p.NatureNum, 2));

            var spd = ruleSet.GetStat(basePkmn.BaseStats[StatTypes.nbSpeed], p.DVs[StatTypes.nbSpeed],
                p.EVs[StatTypes.nbSpeed], p.Level, GetNatureMod(p.NatureNum, 3));

            var spatk = ruleSet.GetStat(basePkmn.BaseStats[StatTypes.nbSpAttack], p.DVs[StatTypes.nbSpAttack],
                p.EVs[StatTypes.nbSpAttack], p.Level, GetNatureMod(p.NatureNum, 4));

            var spdef = ruleSet.GetStat(basePkmn.BaseStats[StatTypes.nbSpDefense], p.DVs[spDefDv],
                p.EVs[StatTypes.nbSpDefense], p.Level, GetNatureMod(p.NatureNum, 5));

            total = (int)Math.Round(hp + atk + def + spd + spatk + spdef); // -- VB rounds to even when assigning to an Integer.

            // -- Applies some type matrix scaling?
            int matrixAdjust = 0;
            for (var i = 1; i < 18; i++)
            {
                float battleDamage;
                
                if (basePkmn.Type2 == Elements.nbNoType)
                    battleDamage = TypeDatabase.BattleMatrix[i, (int)basePkmn.Type1];
                else
                    battleDamage = TypeDatabase.BattleMatrix[i, (int)basePkmn.Type1] * TypeDatabase.BattleMatrix[i, (int)basePkmn.Type2];
                
                switch (battleDamage)
                {
                    case 0:
                        matrixAdjust += 75;
                        break;
                    case 0.25f:
                        matrixAdjust += 50;
                        break;
                    case 0.5f:
                        matrixAdjust += 25;
                        break;
                    case 2:
                        matrixAdjust -= 35;
                        break;
                    case 4:
                        matrixAdjust -= 75;
                        break;
                }
            }
            total += matrixAdjust;
            // -- Skew for OP Pokemon
            switch (p.Pokemon.No)
            {
                case 150: // Mewtwo, Lugia, Ho-Oh
                case 249:
                case 250: 
                case 382: // Kyogyre
                case 383: // Groudon
                case 384: // Raqyuaza
                case 386: // Deoxys
                case 387: // Turtwig... what.. maybe Deoxys subforms?
                case 388: // Grotle
                case 389: // Torterra..
                    total += (p.Level * 5);
                    break;
                // Legendary birds and dogs, mew, Celebi
                case 144 or 145 or 146:
                case 151:
                case 243 or 244 or 245 or 246:
                case 251:
                case 379 or 380 or 381:
                case 385:
                    total = (int)Math.Round(total + p.Level * 2.5);
                    break;
                // Snorlax, Dragonite, Tyranitar.
                case 143:
                case 149:
                case 248:
                    total = (int)Math.Round(total + p.Level * 1.25);
                    break;
            }
            
            return total;
        }
        
        /*
         * Generates a 0-100 power rating for your team
         * Takes all pokemon stats, adds them, then it is turned into a percentage of the calculated maximum for a team.
         * Use the OP Pokemon above to really boost it..
         */
        public static int GetTeamRank(TeamMember[] team)
        {
            if (team[0] == null) {
                return 0;
            }

            int total = team.Sum(GetPokeRank);
            total -= Constants.LowestRank;

            int range = BattleSystem.IsAdvance(team[0].GameVersion)
                ? Constants.ADVHighestRank - Constants.ADVLowestRank
                : Constants.HighestRank - Constants.LowestRank;

            return Math.Min((int)Math.Floor(total * 100.0 / range), 100);
        }

        public static int GetNatureMod(byte workingPokeNatureNum, int stat) {
            if (workingPokeNatureNum >= BattleSystem.NatureStats.Length) {
                return 0;
            }
            
            return BattleSystem.NatureStats[workingPokeNatureNum].StatChg[stat];
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetVolumeInformation(string rootPathName, StringBuilder? volumeNameBuffer, int volumeNameSize,
            out uint volumeSerialNumber, IntPtr maximumComponentLength, IntPtr fileSystemFlags, StringBuilder? fileSystemNameBuffer, int fileSystemNameSize);

        /// <summary>
        /// A recreation of Code.bas GetSerialNumber: the compressed, 13-byte Station ID. Pass it to DecompressSid for the
        /// 21-character form. It hashes the hard drive serial (from SerialGrabber), or the Windows drive's volume serial
        /// when no usable drive serial is found.
        /// </summary>
        /// <param name="useThisOne">Hash this instead of reading the drive. Used as given, without upper-casing (as in the VB).</param>
        public static byte[] GetSerialNumber(string? useThisOne = null) {
            string build = !string.IsNullOrEmpty(useThisOne)
                ? useThisOne
                : OperatingSystem.IsWindows() ? new SerialGrabber().GetSerial().ToUpperInvariant() : "";

            // -- Keep '0' through 'Z' only.
            var temp = new string(build.Where(c => c >= 48 && c <= 90).ToArray());

            if (temp.Length == 0 || temp.All(c => c == '0')) {
                // -- "Bad SID": fall back to the volume serial of the drive Windows is on.
                uint volumeSerial = 0;
                if (OperatingSystem.IsWindows()) {
                    string root = Path.GetPathRoot(Environment.GetFolderPath(Environment.SpecialFolder.Windows)) ?? @"C:\";
                    GetVolumeInformation(root, null, 0, out volumeSerial, IntPtr.Zero, IntPtr.Zero, null, 0);
                }

                temp = volumeSerial.ToString("X"); // -- VB Hex(): uppercase, no padding.
            }

            // -- MD5 hex, each digit as 4 bits, cut to 100 bits: the first 100 bits of the digest.
            byte[] hash = System.Security.Cryptography.MD5.HashData(Encoding.Latin1.GetBytes(temp));
            return BinaryToBytes(BytesToBinary(hash).Substring(0, 100));
        }
    }
}
