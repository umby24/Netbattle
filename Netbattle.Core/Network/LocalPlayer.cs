using Netbattle.Common;
using Netbattle.Core.Common;
using Netbattle.Core.Database;

namespace Netbattle.Core.Network {
    /// <summary>
    /// NbClient.You. The profile fields read straight from UserSettings and Configuration on every access,
    /// so what gets sent to the server always matches the saved settings. Change those settings, not this object.
    /// Server-assigned fields (Id, Wins, Authority, Battling, ...) behave as on any Player.
    /// </summary>
    public class LocalPlayer : Player {
        private static UserSettings Profile => UserSettings.CurrentSettings
            ?? throw new InvalidOperationException("UserSettings must be loaded before connecting.");

        public override string Name {
            get => Profile.Username;
            set => ReadOnly(nameof(Name));
        }

        public override int Picture {
            get => Profile.IconUsed;
            set => ReadOnly(nameof(Picture));
        }

        public override string Description {
            get => Profile.MoreInfo;
            set => ReadOnly(nameof(Description));
        }

        public override List<TeamMember> Team {
            get => Profile.Team.ToList();
            set => ReadOnly(nameof(Team));
        }

        public override int TeamPower {
            get => NbMethods.GetTeamRank(Profile.Team);
            set => ReadOnly(nameof(TeamPower));
        }

        public override int GameVersion {
            get => (int)Profile.CurrentCompatibilityMode;
            set => ReadOnly(nameof(GameVersion));
        }

        public override int GraphicsVersion {
            get => (int)Profile.CurrentGraphicsMode;
            set => ReadOnly(nameof(GraphicsVersion));
        }

        public override int Compatibility {
            get {
                bool[] modes = BattleSystem.CompatCheck(Profile.Team);
                return NbMethods.MakeBinArray(ref modes);
            }
            set => ReadOnly(nameof(Compatibility));
        }

        public override bool ShowTeam {
            get => Configuration.CurrentSettings.ShowTeam;
            set => ReadOnly(nameof(ShowTeam));
        }

        private static void ReadOnly(string property) =>
            throw new InvalidOperationException(
                $"You.{property} comes from the saved settings. Change UserSettings or Configuration instead.");
    }
}
