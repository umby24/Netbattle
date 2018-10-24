using System;
using System.Collections.Generic;
using System.IO;
using Netbattle.Common;
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

        public Configuration() {
            EnableLogging = false;
            LastJnb = "";
            LastPnb = "";
            RegistryIp = "192.168.1.66";
            AdvConnects = new List<string>();
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
