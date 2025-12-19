using System;
using System.IO;
using Netbattle.Common;
using Newtonsoft.Json;

namespace Netbattle.Database {
    public class UserSettings {
        public static UserSettings CurrentSettings;
        public CompatModes CurrentCompatibilityMode { get; set; }
        public GraphicsMode CurrentGraphicsMode { get; set; }
        public string Username { get; set; }
        public string MoreInfo { get; set; }
        public string WinMessage { get; set; }
        public string LoseMessage { get; set; }
        public byte IconUsed { get; set; }
        public Pokemon[] Team;
        public string StationID { get; set; }
        public string DbModName { get; set; }
        public string DbMods { get; set; }
        public string RegistryAddress { get; set; }
        
        public static void Load() {
            if (!string.IsNullOrWhiteSpace(Configuration.CurrentSettings.LastJnb)) {
                LoadJnb();
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
                Team = new Pokemon[6],
                RegistryAddress = "registry.pmnb.net"
            };
        }

        protected static void LoadJnb() {
            try {
                CurrentSettings =
                    JsonConvert.DeserializeObject<UserSettings>(
                        File.ReadAllText(Configuration.CurrentSettings.LastJnb));
                CurrentSettings.CurrentCompatibilityMode = CompatModes.nbFullAdvance;
                for (var i = 0; i < 6; i++) {
                    if (CurrentSettings.Team[i] != null) {
                        CurrentSettings.Team[i] = CurrentSettings.Team[i].SetupFromDatabase();
                        CurrentSettings.Team[i].GameVersion = CompatModes.nbFullAdvance;
                    }
                }
            }
            catch (Exception ex) {
                Logger.Log(LogType.Error, "Error loading JNB File:");
                Logger.Log(ex);
                MakeNewSettings();
            }

            if (string.IsNullOrEmpty(CurrentSettings.RegistryAddress))
            {
                CurrentSettings.RegistryAddress = "registry.pmnb.net";
            }
            if (CurrentSettings.Team.Length != 6)
            {
                var newTeam = new Pokemon[6];
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
            Configuration.CurrentSettings.LastJnb = "team.jnb";
            Configuration.Save();
        }

        protected void SavePnb() {

        }
    }
}
