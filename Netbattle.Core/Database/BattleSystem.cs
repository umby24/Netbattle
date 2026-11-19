using Netbattle.Core.Battle;
using Netbattle.Core.Common;

namespace Netbattle.Core.Database {
    public class BattleSystem {
        public static NatureType[] NatureStats = BuildStatTable();

        /// <summary>
        /// Determine if the current team is eligible to battle.
        /// </summary>
        /// <returns></returns>
        public static bool BattleEligible() {
            if (string.IsNullOrWhiteSpace(UserSettings.CurrentSettings.Username)) {
                return false;
            }

            foreach (var tm in UserSettings.CurrentSettings.Team) {
                if (tm?.Pokemon == null || tm.Pokemon.No == 0) {
                    return false;
                }
                if (tm.Move[0] == 0 || tm.Move[1] == 0 || tm.Move[2] == 0 || tm.Move[3] == 0) {
                    return false;
                }
            }
            
            return true;
        }
        public static int GetGenerationStat(CompatModes mode, int baseVal, int iv, int ev, int level, int pmod) {
            switch (mode) {
                case CompatModes.nbTrueRBY:
                case CompatModes.nbRBYTrade:
                    IGenerationRuleset rbyRules = new RBYRules();
                    return rbyRules.GetStat(baseVal, iv, ev, level, pmod);
                case CompatModes.nbTrueGSC:
                case CompatModes.nbGSCTrade:
                    IGenerationRuleset gscRules = new GSCRules();
                    return gscRules.GetStat(baseVal, iv, ev, level, pmod);
                case CompatModes.nbFullAdvance:
                case CompatModes.nbModAdv:
                case CompatModes.nbTrueRuSa:
                    IGenerationRuleset advRules = new ADVRules();
                    return advRules.GetStat(baseVal, iv, ev, level, pmod);
                default:
                    throw new ArgumentOutOfRangeException(nameof(mode), mode, null);
            }
        }

        public static int GetGenerationHp(CompatModes mode, int baseVal, int iv, int ev, int level) {
            switch (mode) {
                case CompatModes.nbTrueRBY:
                case CompatModes.nbRBYTrade:
                    IGenerationRuleset rbyRules = new RBYRules();
                    return rbyRules.GetHp(baseVal, iv, ev, level);
                case CompatModes.nbTrueGSC:
                case CompatModes.nbGSCTrade:
                    IGenerationRuleset gscRules = new GSCRules();
                    return gscRules.GetHp(baseVal, iv, ev, level);
                case CompatModes.nbFullAdvance:
                case CompatModes.nbModAdv:
                case CompatModes.nbTrueRuSa:
                    IGenerationRuleset advRules = new ADVRules();
                    return advRules.GetHp(baseVal, iv, ev, level);
                default:
                    throw new ArgumentOutOfRangeException(nameof(mode), mode, null);
            }
        }

        /// <summary>
        /// Port of CompatCheck: which of the 7 modes the team is legal in. Every mode starts allowed, and each member can only turn modes off.
        /// Index: 0 RBY (Trades), 1 GSC (Trades), 2 Advance, 3 Advance +, 4 Advance (Mods), 5 True RBY, 6 True GSC.
        /// </summary>
        public static bool[] CompatCheck(IEnumerable<TeamMember?> team) {
            bool[] validModes = [true, true, true, true, true, true, true];
            var rbyRules = new RBYRules();
            var gscRules = new GSCRules();
            var advRules = new ADVRules();
            var advPlusRules = new ADVPlusRules();

            foreach (TeamMember? member in team) {
                if (member?.Pokemon == null || member.Pokemon.No <= 0) continue;

                PokemonSpecies species = member.Pokemon;
                if (!species.ExistRBY) {
                    validModes[0] = false;
                    validModes[5] = false;
                }
                if (!species.ExistGSC) validModes[6] = false;
                if (!(species.ExistGSC || species.ExistRBY)) validModes[1] = false;
                if (!species.ExistADV) validModes[2] = false;

                if ((int)member.Item > 41) {
                    validModes[1] = false;
                    validModes[6] = false;
                }
                else if (!member.Item.IsAdvanceItem()) {
                    validModes[2] = false;
                    validModes[3] = false;
                }

                foreach (int mv in member.Move) {
                    if (mv == 0) continue;
                    Move move = MoveDatabase.Moves[mv];

                    if (!move.RBYMove) {
                        validModes[0] = false;
                        validModes[5] = false;
                    }

                    if (!move.GSCMove) {
                        validModes[1] = false;
                        validModes[6] = false;
                    }

                    if (!rbyRules.CheckMove(species, move)) validModes[5] = false;     // -- TrueRBYCheck
                    if (!gscRules.CheckMove(species, move)) validModes[6] = false;     // -- TrueGSCCheck
                    if (!advRules.CheckMove(species, move)) validModes[2] = false;     // -- TrueAdvCheck
                    if (!advPlusRules.CheckMove(species, move)) validModes[3] = false; // -- TrueAdvPlusCheck
                }
            }

            return validModes;
        }
        
        public static IGenerationRuleset GetGenerationRuleset(CompatModes mode)
        {
            switch (mode)
            {
                case CompatModes.nbTrueRBY:
                case CompatModes.nbRBYTrade:
                    return new RBYRules();
                case CompatModes.nbTrueGSC:
                case CompatModes.nbGSCTrade:
                    return new GSCRules();
                case CompatModes.nbFullAdvance:
                case CompatModes.nbModAdv:
                case CompatModes.nbTrueRuSa:
                    return new ADVRules();
                default:
                    throw new ArgumentOutOfRangeException(nameof(mode), mode, null);
            }
        }
        public static bool IsAdvance(CompatModes mode) {
            return mode is CompatModes.nbTrueRuSa or CompatModes.nbFullAdvance or CompatModes.nbModAdv;
        }

        public static bool ShinyDV(short DVAtk) {
            var vals = new[] { 15, 14, 11, 10, 7, 6, 3, 2 };
            return vals.Contains(DVAtk);
        }
        
        public static void GenerateStatTable() {
            NatureStats = BuildStatTable();
        }

        private static NatureType[] BuildStatTable() {
            var natures = new NatureType[25];

            for (int i = 0; i < 25; i++) {
                natures[i] = new NatureType();

                int y = i / 5 + 1;
                natures[i].StatChg[y] += 1;
                y = i % 5 + 1;
                natures[i].StatChg[y] -= 1;
            }

            return natures;
        }
        
        /// <summary>
        /// Checks if a move is exclusively learned through specified methods
        /// </summary>
        private bool LegalMoveCheck(bool[,] allMoves, short moveNum, MoveTypes c1, MoveTypes c2 = 0, 
            MoveTypes c3 = 0, MoveTypes c4 = 0, MoveTypes c5 = 0, MoveTypes c6 = 0, 
            MoveTypes c7 = 0, MoveTypes c8 = 0, MoveTypes c9 = 0, bool ignoreRBY = false)
        {
            bool build = false;

            try
            {
                // Check if the move exists in any of the specified categories
                if (allMoves[moveNum, (int)c1]) build = true;
                if (c2 > 0 && allMoves[moveNum, (int)c2]) build = true;
                if (c3 > 0 && allMoves[moveNum, (int)c3]) build = true;
                if (c4 > 0 && allMoves[moveNum, (int)c4]) build = true;
                if (c5 > 0 && allMoves[moveNum, (int)c5]) build = true;
                if (c6 > 0 && allMoves[moveNum, (int)c6]) build = true;
                if (c7 > 0 && allMoves[moveNum, (int)c7]) build = true;
                if (c8 > 0 && allMoves[moveNum, (int)c8]) build = true;
                if (c9 > 0 && allMoves[moveNum, (int)c9]) build = true;

                // Check that the move is NOT in any other categories (exclusivity check)
                for (int x = 1; x <= 13; x++)
                {
                    if (x != (int)c1 && x != (int)c2 && x != (int)c3 && x != (int)c4 && 
                        x != (int)c5 && x != (int)c6 && x != (int)c7 && x != (int)c8 && x != (int)c9)
                    {
                        if (allMoves[moveNum, x])
                        {
                            // If ignoreRBY is true, don't fail on RBY moves
                            if (!(ignoreRBY && (x == (int)MoveTypes.nbRBYLevel || x == (int)MoveTypes.nbRBYTM)))
                            {
                                build = false;
                                break;
                            }
                        }
                    }
                }
            }
            catch
            {
                // Error handling - return false on any exception
                build = false;
            }

            return build;
        }
    }
}