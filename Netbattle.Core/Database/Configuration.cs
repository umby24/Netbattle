using Netbattle.Core.Common;
using Newtonsoft.Json;

namespace Netbattle.Core.Database {
    /// <summary>
    /// App-wide options, saved to Configuration.json. VB kept these in the registry under "NetBattle\Options".
    /// Anything that belongs to a team or trainer profile goes in UserSettings instead, so loading a team never changes these.
    /// </summary>
    public class Configuration {
        private const string ConfigFileName = "Configuration.json";
        public const string DefaultRegistryAddress = "registry.pmnb.net";
        public static Configuration CurrentSettings;

        public bool EnableLogging { get; set; }
        public string LastJnb { get; set; }
        public string LastPnb { get; set; }
        public string RegistryAddress { get; set; }
        public List<string> AdvConnects { get; set; }

        // -- The challenge window's last used rules and mode. VB kept these in the registry under "Battle Rules".
        public List<bool> BattleRules { get; set; } = new List<bool>(); // -- Index 0 is rule 1.
        public int BattleMode { get; set; } = 2; // -- Mode list index; VB's default is 2, Advance.
        public bool[] MessageToggles = new bool[10]; // -- VB's MsgToggle, also sent to the server in the USER packet.
        public bool UseTimestamp { get; set; }
        public bool ParseUrls { get; set; }
        public bool EnableBackground { get; set; }
        public bool AnimateHpBars { get; set; }
        public bool Sounds { get; set; }
        public bool Music { get; set; }
        public bool PlayerNamePrefix { get; set; }
        public bool ExtendedMessages { get; set; }
        public bool UseNicknames { get; set; }
        public bool ShowTeam { get; set; } // -- VB's AllowViewing.
        public bool LogsPrompt { get; set; }
        public bool LogsAutosave { get; set; }
        public bool ReplayPrompt { get; set; }
        public bool ReplayAutosave { get; set; }
        public string TimestampFormat { get; set; }
        public string StoredPassword { get; set; } // -- VB's SavedPassword: the MD5 of the user password, as 32 hex characters.

        public Configuration() {
            EnableLogging = false;
            LastJnb = "";
            LastPnb = "";
            RegistryAddress = DefaultRegistryAddress;
            AdvConnects = new List<string>();
            MessageToggles = new[] { true, true, true, true, true, true, true, true, true, true };
            ShowTeam = true; // -- VB defaults "Allow Viewing" to 1.
            ExtendedMessages = false;
            UseNicknames = true;
            Music = true;
            Sounds = true;
            PlayerNamePrefix = true;
            ReplayPrompt = true;
            AnimateHpBars = true;
            EnableBackground = true;
            StoredPassword = "";
        }

        public static void Load() {
            if (!File.Exists(ConfigFileName)) {
                CurrentSettings = new Configuration();
                Save();
                return;
            }

            try {
                CurrentSettings = JsonConvert.DeserializeObject<Configuration>(File.ReadAllText(ConfigFileName));
            }
            catch (Exception ex) {
                Logger.Log(LogType.Error, "Error occured loading configuration file.");
                Logger.Log(ex);
                CurrentSettings = new Configuration();
            }

            if (string.IsNullOrWhiteSpace(CurrentSettings.RegistryAddress))
                CurrentSettings.RegistryAddress = DefaultRegistryAddress;

            CurrentSettings.StoredPassword ??= "";
        }

        public static void Save() {
            try {
                File.WriteAllText(ConfigFileName, JsonConvert.SerializeObject(CurrentSettings, Formatting.Indented));
            }
            catch (Exception ex) {
                Logger.Log(LogType.Error, "Error occured saving configuration file.");
                Logger.Log(ex);
            }
        }
    }
}
