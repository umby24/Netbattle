using System.Collections.Generic;
using System.IO;
using System.Linq;
using Netbattle.Common;

namespace Netbattle.Database {
    public class PbxFile {
        public string FilePath { get; set; }
        public const string FileHeader = "BOX2.1";
        public List<Pokemon> BoxPokemon { get; set; }
        
        public PbxFile(string filePath) {
            FilePath = filePath;
        }

        public void Load() {
            using (var sr = new StreamReader(FilePath)) {
                char[] headerChars = new char[FileHeader.Length];
                sr.Read(headerChars, 0, FileHeader.Length);
                string header = new string(headerChars);
                if (header != FileHeader) {
                    throw new InvalidDataException("Invalid PBX file header.");
                }
                
                BoxPokemon = new List<Pokemon>();
                while (!sr.EndOfStream) {
                    var markerBytes = new char[2];
                    sr.Read(markerBytes, 0, 2);
                    
                    var packedPkmnData = new char[20];
                    sr.Read(packedPkmnData, 0, 20);
                    
                    string decomposed = NbMethods.BytesToBinary(packedPkmnData.Select(c => (byte)c).ToArray());
                    var pkmnObj = Pokemon.FromBinary(decomposed, "");
                    
                    int markerNum = int.Parse(new string(markerBytes), System.Globalization.NumberStyles.HexNumber);
                    pkmnObj.MarkerNum = (byte)markerNum;
                    BoxPokemon.Add(pkmnObj);
                }
                
            }
        }
        
        public void Save() {
            using (StreamWriter sw = new StreamWriter(FilePath)) {
                sw.Write(FileHeader);
                foreach (var pkmn in BoxPokemon) {
                    sw.Write(pkmn.MarkerNum.ToString("X"));
                    sw.WriteLine(pkmn.ToStringBytes());
                }
                sw.Flush();
                sw.Close();
            }
        }
        
    }
}