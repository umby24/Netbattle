using System;
using System.Collections.Generic;
using System.IO;
using Netbattle.Common;
using Netbattle.Core.Common;
using Newtonsoft.Json;

namespace Netbattle {
    public class Configuration {
        private const string ConfigFileName = "Configuration.json";
        public static Configuration CurrentSettings;

        public bool EnableLogging { get; set; }
        public string LastJnb { get; set; }
        public string LastPnb { get; set; }
        public string RegistryIp { get; set; }
        public List<string> AdvConnects { get; set; }

        // -- The challenge window's last used rules and mode. VB kept these in the registry under "Battle Rules".
        public List<bool> BattleRules { get; set; } = new List<bool>(); // -- Index 0 is rule 1.
        public int BattleMode { get; set; } = 2; // -- Mode list index; VB's default is 2, Advance.
        public bool[] MessageToggles = new bool[10];
        public bool UseTimestamp { get; set; }
        public bool ParseUrls { get; set; }
        public bool EnableBackground { get; set; }
        public bool AnimateHpBars { get; set; }
        public bool Sounds { get; set; }
        public bool Music { get; set; }
        public bool PlayerNamePrefix { get; set; }
        public bool ExtendedMessages { get; set; }
        public bool UseNicknames { get; set; }
        public bool ShowTeam { get; set; }
        public bool LogsPrompt { get; set; }
        public bool LogsAutosave { get; set; }
        public bool ReplayPrompt { get; set; }
        public bool ReplayAutosave { get; set; }
        public string TimestampFormat { get; set; }
        public string StoredPassword { get; set; }
        

        public Configuration() {
            EnableLogging = false;
            LastJnb = "";
            LastPnb = "";
            RegistryIp = "192.168.1.66";
            AdvConnects = new List<string>();
            MessageToggles = new[] { true, true, true, true, true, true, true, true, true, true };
            ShowTeam = false;
            ExtendedMessages = false;
            UseNicknames = true;
            Music = true;
            Sounds = true;
            PlayerNamePrefix = true;
            ReplayPrompt = true;
            AnimateHpBars = true;
            EnableBackground = true;
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
