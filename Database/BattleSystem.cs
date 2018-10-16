namespace Netbattle.Database {
    public class BattleSystem {
        public static int GetStat(int level, int basePwr, int dv) {
            return (int)(level * (basePwr + dv + 31.9f) / 50) + 5;
        }

        public static int GetHp(int level, int basePwr, int dv) {
            return (int)(level * (basePwr + dv + 31.9f + 50) / 50) + 10;
        }
    }
}
