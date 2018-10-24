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

        public static void Load() {
            if (!string.IsNullOrWhiteSpace(Configuration.CurrentSettings.LastJnb)) {
                LoadJnb();
                return;
            }

            if (!string.IsNullOrWhiteSpace(Configuration.CurrentSettings.LastPnb)) {
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
                CurrentCompatibilityMode = CompatModes.nbFullAdvance,
                CurrentGraphicsMode = GraphicsMode.nbGFXEme,
                IconUsed = 1,
                Team = new Pokemon[6]
            };
        }

        protected static void LoadJnb() {
            try {
                CurrentSettings =
                    JsonConvert.DeserializeObject<UserSettings>(
                        File.ReadAllText(Configuration.CurrentSettings.LastJnb));
            }
            catch (Exception ex) {
                Logger.Log(LogType.Error, "Error loading JNB File:");
                Logger.Log(ex);
                MakeNewSettings();
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
