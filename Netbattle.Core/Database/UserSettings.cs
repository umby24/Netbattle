using Netbattle.Core.Common;
using Netbattle.Core.Files;
using Newtonsoft.Json;

namespace Netbattle.Core.Database {
    /// <summary>
    /// The trainer profile and team, saved to team.jnb: the same data a .pnb holds.
    /// App-wide options such as the stored password and registry address live in Configuration.
    /// </summary>
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

        // -- Older team.jnb files stored the registry address here. It's read so it can move to Configuration, and never written.
        [JsonProperty("RegistryAddress")]
        private string? LegacyRegistryAddress { get; set; }
        public bool ShouldSerializeLegacyRegistryAddress() => false;

        public static void Load(string lastJnb = null, string lastPnb = null) {
            if (!string.IsNullOrWhiteSpace(lastJnb) && File.Exists(lastJnb)) {
                LoadJnb(lastJnb);
                return;
            }

            if (!string.IsNullOrEmpty(lastPnb) && File.Exists(lastPnb)) {
                LoadPnb(lastPnb);
                return;
            }

            MakeNewSettings();
        }

        public static void Save(string filePath = "team.jnb") {
            SaveJnb(filePath);
        }

        protected static void MakeNewSettings() {
            CurrentSettings = new UserSettings {
                CurrentCompatibilityMode = CompatModes.nbModAdv,
                CurrentGraphicsMode = GraphicsMode.nbGFXEme,
                IconUsed = 1,
                Team = new TeamMember[6]
            };
        }

        protected static void LoadJnb(string filePath) {
            CurrentSettings = JsonConvert.DeserializeObject<UserSettings>(File.ReadAllText(filePath));

            var team = new TeamMember[6];
            for (var i = 0; i < Math.Min(CurrentSettings.Team?.Length ?? 0, 6); i++) {
                team[i] = CurrentSettings.Team[i];
            }
            CurrentSettings.Team = team;

            for (var i = 0; i < 6; i++) {
                if (team[i]?.Pokemon == null) {
                    team[i] = null;
                    continue;
                }

                team[i].Initialize(CurrentSettings.CurrentCompatibilityMode);
            }

            MigrateLegacyRegistryAddress();
        }

        // -- Carries a custom registry address from an old team.jnb into Configuration, unless Configuration already has its own.
        private static void MigrateLegacyRegistryAddress() {
            string? legacy = CurrentSettings.LegacyRegistryAddress;
            CurrentSettings.LegacyRegistryAddress = null;
            Configuration config = Configuration.CurrentSettings;

            if (config == null || string.IsNullOrWhiteSpace(legacy) || legacy == Configuration.DefaultRegistryAddress)
                return;

            if (config.RegistryAddress != Configuration.DefaultRegistryAddress)
                return;

            config.RegistryAddress = legacy;
            Configuration.Save();
        }

        protected static void LoadPnb(string filePath) {
            var pnb = new PnbFile(filePath);
            pnb.Load();

            MakeNewSettings();
            CurrentSettings.Username = pnb.Name;
            CurrentSettings.MoreInfo = pnb.ExtraInfo;
            CurrentSettings.WinMessage = pnb.WinMessage;
            CurrentSettings.LoseMessage = pnb.LoseMessage;
            CurrentSettings.IconUsed = pnb.CurrentPicture;
            CurrentSettings.CurrentGraphicsMode = (GraphicsMode)pnb.GraphicsVersion;
            CurrentSettings.DbModName = pnb.DbModName;
            CurrentSettings.DbMods = pnb.DbModString;
            CurrentSettings.Team = pnb.Team;

            TeamMember? first = pnb.Team.FirstOrDefault(a => a != null);
            if (first != null)
                CurrentSettings.CurrentCompatibilityMode = first.GameVersion;
        }

        protected static void SaveJnb(string filePath) {
            File.WriteAllText(filePath, JsonConvert.SerializeObject(CurrentSettings, Formatting.Indented));
        }

        protected void SavePnb() {

        }
    }
}
