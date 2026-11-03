using System.Text;
using Netbattle.Core.Common;

namespace Netbattle.Core.Files {
    public class PbxFile {
        public string FilePath { get; set; }
        private const string FileHeader = "BOX2.1";
        private const string OldFileHeader = "BOX2.0"; // -- Same format, but with 10 character nicknames.
        public List<TeamMember> BoxPokemon { get; set; }

        public PbxFile(string filePath) {
            FilePath = filePath;
        }

        public void Load() {
            byte[] file = System.IO.File.ReadAllBytes(FilePath);
            string header = file.Length >= FileHeader.Length ? Encoding.Latin1.GetString(file, 0, FileHeader.Length) : "";

            if (header != FileHeader && header != OldFileHeader) {
                throw new InvalidDataException("Invalid PBX file header.");
            }

            int nicknameLength = header == OldFileHeader ? 10 : 15;
            int recordLength = 1 + nicknameLength + 20; // -- Marker is one hex character.

            BoxPokemon = [];
            for (int i = FileHeader.Length; i + recordLength <= file.Length; i += recordLength) {
                var marker = Convert.ToByte(Encoding.Latin1.GetString(file, i, 1), 16);
                string nickname = Encoding.Latin1.GetString(file, i + 1, nicknameLength);
                string decomposed = NbMethods.BytesToBinary(file.Skip(i + 1 + nicknameLength).Take(20).ToArray());

                TeamMember? pkmnObj = TeamMember.FromBinary(decomposed, nickname);
                if (pkmnObj == null)
                    continue;

                pkmnObj.Initialize(pkmnObj.GameVersion);
                pkmnObj.MarkerNum = marker;
                BoxPokemon.Add(pkmnObj);
            }
        }

        public void Save() {
            using var fs = new FileStream(FilePath, FileMode.Create);
            fs.Write(Encoding.Latin1.GetBytes(FileHeader));

            foreach (TeamMember pkmn in BoxPokemon) {
                fs.Write(Encoding.Latin1.GetBytes(pkmn.MarkerNum.ToString("X")));
                fs.Write(pkmn.ToStringBytes());
            }
        }

    }
}
