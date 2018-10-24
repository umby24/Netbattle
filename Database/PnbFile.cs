using System;
using System.IO;
using System.Text;
using Netbattle.Common;

namespace Netbattle.Database {
    class PnbFile {
        public string Name { get; set; }
        public string ExtraInfo { get; set; }
        public string WinMessage { get; set; }
        public string LoseMessage { get; set; }
        public byte TBMode { get; set; } // -- ???
        public byte CurrentPicture { get; set; }
        public byte Version { get; set; }
        public Pokemon[] Team { get; set; }

        private const string FileHeader = " PNB4.1";
        private readonly string _filePath;

        public PnbFile(string filePath) {
            _filePath = filePath;
            Team = new Pokemon[6];
        }

        public void Save() {
            // FileHeader + (nameLength as byte) + name
            // -- + (ExtraLength as byte) + Extra
            // -- + (WinLength as byte) + WinMessage
            // -- + (LoseMess Length as byte) + LoseMessage
            // -- + TBMode as byte
            // -- + Chosen Picture
            // -- + Version.
            // -- Uses 'You' type in Code.bas.
            // -- Next, does PKMN 2 STR for each pokemon in the team..
            // -- if Modded DB, temp += DbModName & DbModStr
            // -- Saved as is like that. :)
        }

        public void Load() {
            byte[] myFile = File.ReadAllBytes(_filePath);
            Team = new Pokemon[6];

            using (var ms = new MemoryStream(myFile)) {
                using (var br = new BinaryReader(ms)) {
                    br.ReadBytes(7); // -- Should be the header.
                    byte nameLen = br.ReadByte();
                    Name = Encoding.ASCII.GetString(br.ReadBytes(nameLen));
                    nameLen = br.ReadByte();
                    ExtraInfo = Encoding.ASCII.GetString(br.ReadBytes(nameLen));
                    nameLen = br.ReadByte();
                    WinMessage = Encoding.ASCII.GetString(br.ReadBytes(nameLen));
                    nameLen = br.ReadByte();
                    LoseMessage = Encoding.ASCII.GetString(br.ReadBytes(nameLen));
                    TBMode = br.ReadByte();
                    CurrentPicture = br.ReadByte();
                    Version = br.ReadByte();

                    for (var i = 0; i < 6; i++) { // -- Read all 6 pokemon..
                        string pokemonNickname = Encoding.ASCII.GetString(br.ReadBytes(15));
                        string decomposed = NbMethods.BytesToBinary(br.ReadBytes(20));
                        Team[i] = Pokemon.FromBinary(decomposed, pokemonNickname);
                    }

                    // -- TODO: Implement Database mod support.
                }
            }

            Logger.Log(LogType.Debug, $"PNB File {_filePath} loaded successfully.");
        }
    }
}
