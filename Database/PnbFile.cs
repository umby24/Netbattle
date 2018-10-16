namespace Netbattle.Database {
    class PnbFile {
        private const string FileHeader = " PNB4.1";

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
    }
}
