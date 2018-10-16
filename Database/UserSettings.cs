using Netbattle.Common;

namespace Netbattle.Database {
    public class UserSettings {
        public static CompatModes CurrentCompatibilityMode { get; set; }
        public static GraphicsMode CurrentGraphicsMode { get; set; }
        public static string Username { get; set; }
        public static string MoreInfo { get; set; }
        public static string WinMessage { get; set; }
        public static string LoseMessage { get; set; }
        public static byte IconUsed { get; set; }
        public static Pokemon[] Team;
        public static string StationID { get; set; }
    }
}
