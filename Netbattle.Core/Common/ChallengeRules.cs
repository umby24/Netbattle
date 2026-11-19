using System.Globalization;

namespace Netbattle.Core.Common {
    // -- Battle rules, as in Code.bas's RuleEnum. 1-based, because the wire format numbers them from bit 0 = rule 1.
    public enum RuleEnum {
        nbSleep = 1,
        nbFreeze,
        nbSelfKO,
        nbUsePPUps,
        nbStadiumMode,
        nbRandbat,
        nbLevelBalance,
        nbTimeout,
        nbUnrated,
        nbExactHP,
        nbNoWatch,
        nbPresentRule
    }

    /// <summary>
    /// The settings of a challenge, as sent in CHLN: and PACC:.
    /// </summary>
    public struct ChallengeSettings {
        public int Mode { get; set; }    // -- 1 RBY, 2 GSC, 3 Advance, 4 Advance Double Battle.
        public int Terrain { get; set; } // -- 0 for random, otherwise (int)TerrainTypes + 1.
        public bool[] Rules { get; set; } // -- Indexed by RuleEnum; index 0 is unused.
    }

    public static class ChallengeRules {
        public const int RuleCount = 15; // -- RuleSelected(1 To 15)

        public static readonly string[] ModeText = {
            "RBY Mode",
            "GSC Mode",
            "Advance Mode",
            "Advance Double Battle Mode"
        };

        private static readonly Dictionary<RuleEnum, string> RuleTextMap = new() {
            [RuleEnum.nbSleep] = "Sleep Clause",
            [RuleEnum.nbFreeze] = "Freeze Clause",
            [RuleEnum.nbSelfKO] = "Self-KO Clause",
            [RuleEnum.nbUsePPUps] = "Apply PP Ups",
            [RuleEnum.nbStadiumMode] = "Stadium Mode",
            [RuleEnum.nbLevelBalance] = "Level Balance",
            [RuleEnum.nbNoWatch] = "Disallow Spectators",
            [RuleEnum.nbTimeout] = "Battle Timeout",
            [RuleEnum.nbUnrated] = "Unrated",
            [RuleEnum.nbPresentRule] = "Use Stadium Present",
            [RuleEnum.nbExactHP] = "Show Enemy HP",
            [RuleEnum.nbRandbat] = "Challenge Cup",
        };

        // -- From Code.bas. VB writes line breaks as %n.
        private static readonly Dictionary<RuleEnum, string> RuleToolTipMap = new() {
            [RuleEnum.nbSleep] = "With this rule enabled, only one Pokémon can be Asleep %nper team.  If a move is used that would violate this%nrule, it will automatically fail.  Pokémon under the%neffects of the move Rest do not count for this rule.",
            [RuleEnum.nbFreeze] = "With this rule enabled, only one Pokémon can be Frozen %nper team.",
            [RuleEnum.nbSelfKO] = "With this rule enabled, if your team has only one Pokémon%nremaining, the moves Destiny Bond and Perish Song will%nalways fail when used by that Pokémon.  Also, if your final%nPokémon uses SelfDestruct or Explosion to Knock Out your%nopponent's final Pokémon, you will automatically lose.  The%npurpose of this rule is to prevent the match from ending in a%nTie.",
            [RuleEnum.nbUsePPUps] = "With this rule enabled, all moves on both your and your%nopponent's Pokémon will have 60% more Power Points.",
            [RuleEnum.nbStadiumMode] = "With this rule enabled, the match will be 3 on 3.  Each player%nwill be allowed to pick three of their six Pokémon to use in%nbattle.  The remaining three Pokémon will not be used.",
            [RuleEnum.nbRandbat] = "With this rule enabled, each player will use six random%nPokémon instead of his/her own team.  The moves of each%nPokémon will also be random, however, each one will always%nhave at least one damaging attack.",
            [RuleEnum.nbLevelBalance] = "With this rule enabled, the Levels of each Pokémon will be%naltered to balance its power.  Pokémon with unusually high%nstats will be assigned a lower level, and Pokémon with lower%nstats will be assigned a higher level.",
            [RuleEnum.nbNoWatch] = "With this rule enabled, spectators will be unable to watch your%nbattle.  However, mods and admins will still be able to watch for%nmoderation purposes.",
            [RuleEnum.nbExactHP] = "With this rule enabled, both players will see the exact HP of%nthe opposing Pokémon instead of a percentage.",
            [RuleEnum.nbTimeout] = "With this rule enabled, after one player has chosen a move, a%ntimer will start.  After 5 minutes of inactivity, the battle will%nautomatically end and the active player will be awarded the%nwin.",
            [RuleEnum.nbUnrated] = "With this rule enabled, the outcome of the battle will affect%nneither players' battle record.",
            [RuleEnum.nbPresentRule] = "With this rule enabled, the formulas used to calculate the%ndamage for the move Present will be those of Pokémon%nCrystal and Pokémon Stadium 2.  If not enabled, the formulas%nin Pokémon Gold and Pokémon Silver will be used.",
        };

        private static readonly Dictionary<TerrainTypes, string> TerrainTextMap = new() {
            [TerrainTypes.nbStadium] = "Stadium",
            [TerrainTypes.nbShortGrass] = "Short Grass",
            [TerrainTypes.nbTallGrass] = "Tall Grass",
            [TerrainTypes.nbVeryTallGrass] = "Very Tall Grass",
            [TerrainTypes.nbOcean] = "Ocean",
            [TerrainTypes.nbUnderwater] = "Underwater",
            [TerrainTypes.nbPond] = "Pond",
            [TerrainTypes.nbsAnd] = "Sand",
            [TerrainTypes.nbCave] = "Cave",
            [TerrainTypes.nbMountain] = "Rock",
        };

        public static string RuleText(int rule) =>
            RuleTextMap.TryGetValue((RuleEnum)rule, out string? text) ? text : "";

        public static string RuleToolTip(int rule) =>
            RuleToolTipMap.TryGetValue((RuleEnum)rule, out string? text) ? text.Replace("%n", Environment.NewLine) : "";

        public static string TerrainText(TerrainTypes terrain) =>
            TerrainTextMap.TryGetValue(terrain, out string? text) ? text : "";

        /// <summary>
        /// Port of CompatVersion: which battle engine a compatibility mode plays under.
        /// </summary>
        public static BattleModes CompatVersion(int compat) => compat switch {
            0 or 5 => BattleModes.nbRBYBattle,
            1 or 6 => BattleModes.nbGSCBattle,
            _ => BattleModes.nbAdvBattle // -- 2, 3, 4, 7
        };

        /// <summary>
        /// Port of MakeBinArray: rule N becomes bit N-1.
        /// </summary>
        public static long MakeBinArray(bool[] rules) {
            long value = 0;
            for (var i = 1; i < rules.Length && i <= RuleCount; i++) {
                if (rules[i])
                    value |= 1L << (i - 1);
            }
            return value;
        }

        /// <summary>
        /// Port of ReadBinArray, for RuleSelected(1 To 15).
        /// </summary>
        public static bool[] ReadBinArray(long value) {
            var rules = new bool[RuleCount + 1];
            for (var i = 1; i <= RuleCount; i++)
                rules[i] = (value & (1L << (i - 1))) != 0;
            return rules;
        }

        // -- VB's FixedHex and Dec.
        public static string FixedHex(long number, int length) {
            string hex = number.ToString("X");
            return hex.Length > length ? hex.Substring(hex.Length - length) : hex.PadLeft(length, '0');
        }

        public static long Dec(string hex) =>
            long.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out long value) ? value : 0;
    }
}
