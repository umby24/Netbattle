using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Netbattle.Common {
    public class CdbFile {
        public string Filename { get; set; }
        public string[] IndividualLines;
        public List<string[]> LineContent;

        public CdbFile(string filePath) {
            if (!File.Exists(filePath) || !filePath.EndsWith(".cdb"))
                throw new InvalidDataException("Invalid CDB file provided."); // -- Invalid 

            Filename = filePath;
        }

        public void Load() {
            LineContent = new List<string[]>();

            byte[] rawFile = File.ReadAllBytes(Filename);
            rawFile = DropDatabaseLength(rawFile);
            
            byte[] decompressed = GZip.Decompress2(rawFile);

            if (decompressed == null) {
                throw new Exception("Failed to decompress CDB File.");
            }

            IndividualLines = Encoding.UTF8.GetString(decompressed).Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string individualLine in IndividualLines) {
                string[] lineSplit = CsvLineSplit(individualLine);
                LineContent.Add(lineSplit);
            }
        }

        public void Save() {

        }

        /// <summary>
        /// Strips out CSV escapes and converts CSV values into something easier to parse.
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        private static string CleanCsvLine(string line) {
            if (line.StartsWith("\"") && line.EndsWith("\""))
                line = line.Substring(1, line.Length - 2);

            if (line == "#TRUE#")
                line = "1";

            if (line == "#FALSE#")
                line = "0";

            return line;
        }

        /// <summary>
        /// Parses out CSV values, including with their respective escaped characters and all.
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        private static string[] CsvLineSplit(string line) {
            var myResult = new List<string>();
            var build = "";
            var inQuote = false;

            foreach (char mChar in line) {
                switch (mChar) {
                    case ',' when !inQuote:
                        myResult.Add(CleanCsvLine(build));
                        build = "";
                        continue;
                    case '"' when inQuote:
                        inQuote = false;
                        continue;
                    case '"':
                        inQuote = true;
                        continue;
                    default:
                        build += mChar;
                        break;
                }
            }

            // -- Don't forget the last one :)
            myResult.Add(CleanCsvLine(build));
            return myResult.ToArray();
        }

        /// <summary>
        /// Helper method to remove the uncompressed length + CRLF in front of .cdb files
        /// also strips the gzip header for you.
        /// </summary>
        /// <param name="input">cdb byte array</param>
        /// <returns>cdb byte array ready to be decompressed</returns>
        private static byte[] DropDatabaseLength(byte[] input) {
            byte[] result = null;

            for (var i = 0; i < input.Length; i++) {
                if (input[i] != 10 || input[i - 1] != 13)
                    continue;

                result = new byte[input.Length - (i + 3)];
                Buffer.BlockCopy(input, i + 3, result, 0, result.Length);
                break;
            }

            return result;
        }
    }
}
