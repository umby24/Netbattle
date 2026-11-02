using Netbattle.Core.Common;
using Newtonsoft.Json;

namespace Netbattle.Core.Database {
    public class UserSettings {
        public static UserSettings? CurrentSettings;
        public CompatModes CurrentCompatibilityMode { get; set; }
        public GraphicsMode CurrentGraphicsMode { get; set; }
        public string Username { get; set; }
        public string MoreInfo { get; set; }
        public string WinMessage { get; set; }
        public string LoseMessage { get; set; }
        public byte IconUsed { get; set; }
        public TeamMember[] Team;
        public string StationID { get; set; }
        public string DbModName { get; set; }
        public string DbMods { get; set; }
        public string RegistryAddress { get; set; }
        
        public static void Load(string lastJnb = null, string lastPnb = null) {
            if (!string.IsNullOrWhiteSpace(lastJnb)) {
                LoadJnb();
                return;
            }

            if (!string.IsNullOrEmpty(lastPnb)) {
                LoadPnb();
                return;
            }
            
            MakeNewSettings();
        }

        public static void Save() {
            SaveJnb();
        }

        protected static void MakeNewSettings() {
            CurrentSettings = new UserSettings {
                CurrentCompatibilityMode = CompatModes.nbModAdv,
                CurrentGraphicsMode = GraphicsMode.nbGFXEme,
                IconUsed = 1,
                Team = new TeamMember[6],
                RegistryAddress = "registry.pmnb.net"
            };
        }

        protected static void LoadJnb(string filePath = "team.jnb") {

                CurrentSettings =
                    JsonConvert.DeserializeObject<UserSettings>(
                        File.ReadAllText(filePath));
                
                for (var i = 0; i < 6; i++) {
                    if (CurrentSettings.Team[i] != null) {
                        CurrentSettings.Team[i].Initialize(CurrentSettings.CurrentCompatibilityMode);
                        CurrentSettings.Team[i].GameVersion = CurrentSettings.CurrentCompatibilityMode;
                    }
                }


            if (string.IsNullOrEmpty(CurrentSettings.RegistryAddress))
            {
                CurrentSettings.RegistryAddress = "registry.pmnb.net";
            }
            if (CurrentSettings.Team.Length != 6)
            {
                var newTeam = new TeamMember[6];
                for (var i = 0; i < Math.Min(CurrentSettings.Team.Length, 6); i++)
                {
                    newTeam[i] = CurrentSettings.Team[i];
                }
                CurrentSettings.Team = newTeam;
            }
        }

        protected static void LoadPnb() {

        }

        protected static void SaveJnb() {
            File.WriteAllText("team.jnb", JsonConvert.SerializeObject(CurrentSettings, Formatting.Indented));
        }

        protected void SavePnb() {

        }
    }
}
