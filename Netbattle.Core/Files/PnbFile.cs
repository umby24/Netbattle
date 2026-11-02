using System.Text;
using Netbattle.Core.Common;

namespace Netbattle.Core.Files {
    public class PnbFile {
        public string Name { get; set; }
        public string ExtraInfo { get; set; }
        public string WinMessage { get; set; }
        public string LoseMessage { get; set; }
        public byte TBMode { get; set; } // -- ???
        public byte CurrentPicture { get; set; }
        public byte GraphicsVersion { get; set; }
        public TeamMember[] Team { get; set; }
        public string DbModName { get; set; }
        public string DbModString { get; set; }
        
        private const string FileHeader = " PNB4.1";
        private readonly string _filePath;

        public PnbFile(string filePath) {
            _filePath = filePath;
            Team = new TeamMember[6];
        }

        public void Save() {
            using (var br = new BinaryWriter(new FileStream(_filePath, FileMode.Create)))
            {
                br.Write(FileHeader);
                br.Write((byte)Name.Length);
                br.Write(Name.ToCharArray());
                br.Write((byte)ExtraInfo.Length);
                br.Write(ExtraInfo.ToCharArray());
                br.Write((byte)WinMessage.Length);
                br.Write(WinMessage.ToCharArray());
                br.Write((byte)LoseMessage.Length);
                br.Write(LoseMessage.ToCharArray());
                br.Write((byte)TBMode);
                br.Write((byte)CurrentPicture);
                br.Write((byte)GraphicsVersion);
                for (var i = 0; i < Team.Length; i++)
                {
                    br.Write(Team[i].ToStringBytes());
                }

                if ((CompatModes)GraphicsVersion == CompatModes.nbModAdv)
                {
                    br.Write(DbModName.PadRight(20));
                    br.Write(DbModString);
                }
            }
        }

        public void Load() {
            byte[] myFile = System.IO.File.ReadAllBytes(_filePath);
            Team = new TeamMember[6];

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
                    GraphicsVersion = br.ReadByte();

                    for (var i = 0; i < 6; i++) { // -- Read all 6 pokemon..
                        string pokemonNickname = Encoding.ASCII.GetString(br.ReadBytes(15));
                        string decomposed = NbMethods.BytesToBinary(br.ReadBytes(20));
                        Team[i] = TeamMember.FromBinary(decomposed, pokemonNickname);
                        Team[i].RegenCurrentStats();
                    }

                    DbModName = Encoding.ASCII.GetString(br.ReadBytes(20)).TrimEnd('\0');
                    DbModString = Encoding.ASCII.GetString(br.ReadBytes((int)(br.BaseStream.Length - br.BaseStream.Position)));
                }
            }

            Logger.Log(LogType.Debug, $"PNB File {_filePath} loaded successfully.");
        }
    }
}
