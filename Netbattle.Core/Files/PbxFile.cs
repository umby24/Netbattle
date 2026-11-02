using Netbattle.Core.Common;

namespace Netbattle.Core.Files {
    public class PbxFile {
        public string FilePath { get; set; }
        private const string FileHeader = "BOX2.1";
        public List<TeamMember> BoxPokemon { get; set; }
        
        public PbxFile(string filePath) {
            FilePath = filePath;
        }

        public void Load() {
            using var sr = new StreamReader(FilePath);
            var headerChars = new char[FileHeader.Length];
            sr.Read(headerChars, 0, FileHeader.Length);
            var header = new string(headerChars);
            if (header != FileHeader) {
                throw new InvalidDataException("Invalid PBX file header.");
            }
                
            BoxPokemon = [];
            while (!sr.EndOfStream) {
                var markerBytes = new char[2];
                sr.Read(markerBytes, 0, 2);
                    
                var packedPkmnData = new char[20];
                sr.Read(packedPkmnData, 0, 20);
                    
                string decomposed = NbMethods.BytesToBinary(packedPkmnData.Select(c => (byte)c).ToArray());
                TeamMember pkmnObj = TeamMember.FromBinary(decomposed, "");
                    
                int markerNum = int.Parse(new string(markerBytes), System.Globalization.NumberStyles.HexNumber);
                pkmnObj.MarkerNum = (byte)markerNum;
                BoxPokemon.Add(pkmnObj);
            }
        }
        
        public void Save() {
            using var sw = new StreamWriter(FilePath);
            sw.Write(FileHeader);
            
            foreach (TeamMember pkmn in BoxPokemon) {
                sw.Write(pkmn.MarkerNum.ToString("X"));
                sw.WriteLine(pkmn.ToStringBytes());
            }
            
            sw.Flush();
            sw.Close();
        }
        
    }
}