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
        private const string OldFileHeader = " PNB4.0"; // -- Same format, but with 10 character nicknames.
        private const int NicknameLength = 15;
        private const int OldNicknameLength = 10;
        private readonly string _filePath;

        public PnbFile(string filePath) {
            _filePath = filePath;
            Team = new TeamMember[6];
        }

        public void Save() {
            using (var br = new BinaryWriter(new FileStream(_filePath, FileMode.Create)))
            {
                br.Write(Encoding.Latin1.GetBytes(FileHeader));
                WritePrefixedString(br, Name);
                WritePrefixedString(br, ExtraInfo);
                WritePrefixedString(br, WinMessage);
                WritePrefixedString(br, LoseMessage);
                br.Write((byte)TBMode);
                br.Write((byte)CurrentPicture);
                br.Write((byte)GraphicsVersion);
                for (var i = 0; i < Team.Length; i++)
                {
                    br.Write(Team[i]?.ToStringBytes() ?? TeamMember.EmptySlotBytes());
                }

                if (Team[0]?.GameVersion == CompatModes.nbModAdv)
                {
                    br.Write(Encoding.Latin1.GetBytes((DbModName ?? "").PadRight(20)));
                    br.Write(Encoding.Latin1.GetBytes(DbModString ?? ""));
                }
            }
        }

        private static void WritePrefixedString(BinaryWriter writer, string value) {
            byte[] data = Encoding.Latin1.GetBytes(value ?? "");
            writer.Write((byte)data.Length);
            writer.Write(data);
        }

        public void Load() {
            byte[] myFile = System.IO.File.ReadAllBytes(_filePath);
            Team = new TeamMember[6];

            using (var ms = new MemoryStream(myFile)) {
                using (var br = new BinaryReader(ms)) {
                    string header = Encoding.Latin1.GetString(br.ReadBytes(7));
                    if (header != FileHeader && header != OldFileHeader)
                        throw new InvalidDataException($"Unsupported PNB file: {_filePath}");

                    byte nameLen = br.ReadByte();
                    Name = Encoding.Latin1.GetString(br.ReadBytes(nameLen));
                    nameLen = br.ReadByte();
                    ExtraInfo = Encoding.Latin1.GetString(br.ReadBytes(nameLen));
                    nameLen = br.ReadByte();
                    WinMessage = Encoding.Latin1.GetString(br.ReadBytes(nameLen));
                    nameLen = br.ReadByte();
                    LoseMessage = Encoding.Latin1.GetString(br.ReadBytes(nameLen));
                    TBMode = br.ReadByte();
                    CurrentPicture = br.ReadByte();
                    GraphicsVersion = br.ReadByte();

                    int nicknameLength = header == OldFileHeader ? OldNicknameLength : NicknameLength;

                    for (var i = 0; i < 6; i++) { // -- Read all 6 pokemon..
                        string pokemonNickname = Encoding.Latin1.GetString(br.ReadBytes(nicknameLength));
                        string decomposed = NbMethods.BytesToBinary(br.ReadBytes(20));
                        Team[i] = TeamMember.FromBinary(decomposed, pokemonNickname);
                        Team[i]?.Initialize(Team[i].GameVersion);
                    }

                    if (Team[0]?.GameVersion == CompatModes.nbModAdv) {
                        DbModName = Encoding.Latin1.GetString(br.ReadBytes(20)).Trim();
                        DbModString = Encoding.Latin1.GetString(br.ReadBytes((int)(br.BaseStream.Length - br.BaseStream.Position)));
                    }
                }
            }

            Logger.Log(LogType.Debug, $"PNB File {_filePath} loaded successfully.");
        }
    }
}
