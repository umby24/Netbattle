using System;
using System.Linq;
using Netbattle.Common;

namespace Netbattle.Database {
    public class BattleSystem {
        public static NatureType[] NatureStats;

        public static int GetStat(int level, int basePwr, int dv) {
            return (int)(level * (basePwr + dv + 31.9f) / 50) + 5;
        }

        public static int GetHp(int level, int basePwr, int dv) {
            return (int)(level * (basePwr + dv + 31.9f + 50) / 50) + 10;
        }

        public static int GetAdvStat(int baseVal, int iv, int ev, int level, int pmod) {
            float p = 1;

            if (pmod == 1) p = 1.1f;
            if (pmod == 0) p = 1.0f;
            if (pmod == -1) p = 0.9f;

            return (int)Math.Round(((baseVal * 2 + iv + (ev / 4)) * level / 100 + 5) * p);
        }

        public static int GetAdvHp(int baseVal, int iv, int ev, int level) {
            if (baseVal == 1) return 1; // Shedninja
            return (int)Math.Round((baseVal * 2.0f + iv + (ev / 4)) * level / 100 + 10 + level);
        }

        public static bool ShinyDV(short DVAtk) {
            var vals = new[] { 15, 14, 11, 10, 7, 6, 3, 2 };
            return vals.Contains(DVAtk);
        }
        
        public static void GenerateStatTable() {
            NatureStats = new NatureType[25];

            for (int i = 0; i < 24; i++) {
                if (NatureStats[i] == null) {
                    NatureStats[i] = new NatureType();
                }

                int y = i / 5 + 1;
                NatureStats[i].StatChg[y] += 1;
                y = i % 5 + 1;
                NatureStats[i].StatChg[y] -= 1;
            }
        }
    }
}