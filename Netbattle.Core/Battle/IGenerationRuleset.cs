using Netbattle.Core.Common;
using Netbattle.Core.Pokemon;

namespace Netbattle.Core.Battle;

public interface IGenerationRuleset
{
    int GetStat(int baseVal, int iv, int ev, int level, int pmod);
    int GetHp(int baseVal, int iv, int ev, int level);
    int HiddenPowerStrength(StatSet DVs);
    Elements HiddenPowerType(StatSet DVs);
}