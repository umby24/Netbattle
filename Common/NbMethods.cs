using System;
using System.Linq;
using System.Text;

namespace Netbattle.Common {
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

        /// <summary>
        /// Convert an integer into a bit string.
        /// </summary>
        /// <param name="input"></param>
        /// <param name="length"></param>
        public static string Dec2Bin(int input, int length) {
            byte[] asBytes = BitConverter.GetBytes(input);
            Array.Reverse(asBytes);
            string asBits = BytesToBinary(asBytes);

            if (asBits.Length < length)
                asBits = asBits.PadLeft(length, '0');
            else {
                asBits = asBits.TrimStart('0').PadLeft(length, '0');
            }

            return asBits;
        }
        /// 000000001  - 1
        /// 000000010 - 2
        /// 000011110 - 30
        /// 001100100 - 100
        

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
            string hexed = Convert.ToString(number, 16); // -- Convert the number to base-16 (HEX)

            if (hexed.Length == length)
                return hexed;

            if (hexed.Length <= length)
                return hexed.PadLeft(length, '0'); // -- length > hexed.length

            int diff = hexed.Length - length;
            hexed = hexed.Substring(diff, hexed.Length - diff);
            return hexed;
        }

        public static string BytesToBinary(byte[] bytes) {
            var build = "";

            foreach (byte b in bytes) {
                string chr = Convert.ToString(b, 2);
                chr = chr.PadLeft(8, '0');
                build += chr;
            }

            return build;
        }

        public static byte[] BinaryToBytes(string binary) {
            var additional = 0;

            if (binary.Length % 8 > 0)
                additional = 1;

            var build = new byte[(binary.Length / 8) + additional];

            var c = 0;
            for (var i = 0; i < binary.Length; i+=8) {
                int len = (binary.Length < i + 8 ? (binary.Length - i) : 8);
                build[c] = Convert.ToByte(binary.Substring(i, len), 2);
                c++;
            }
            
            return build;
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
    }
}
