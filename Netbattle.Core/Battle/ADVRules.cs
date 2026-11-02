using Netbattle.Core.Common;
using Netbattle.Core.Database;
using Netbattle.Core.Pokemon;

namespace Netbattle.Core.Battle;

public class ADVRules : IGenerationRuleset
{
    public int GetStat(int baseVal, int iv, int ev, int level, int pmod) {
        float p = pmod switch {
            1 => 1.1f,
            0 => 1.0f,
            -1 => 0.9f,
            _ => 1
        };

        return (int)Math.Round(((baseVal * 2 + iv + (ev / 4)) * level / 100 + 5) * p);
    }

    public int GetHp(int baseVal, int iv, int ev, int level) {
        if (baseVal == 1) return 1; // Shedninja
        return (int)Math.Round((baseVal * 2.0f + iv + (ev / 4)) * level / 100 + 10 + level);
    }

    public bool CheckMove(PokemonSpecies species, Move move)
    {
        var result = species.Moves.Sets[MoveTypes.nbAdvLevel].Contains(move);
        
        if ( species.Moves.Sets[MoveTypes.nbAdvTM].Contains(move)) result = true;
        if ( species.Moves.Sets[MoveTypes.nbAdvEgg].Contains(move)) result = true;
        if ( species.Moves.Sets[MoveTypes.nbAdvSpecial].Contains(move)) result = true;
        if ( species.Moves.Sets[MoveTypes.nbAdvTutor].Contains(move)) result = true;
        
        return result;
    }

    public int HiddenPower(TeamMember member)
    {
        var result = 0;
        
        if ((member.DVs[StatTypes.nbHP] / 2) % 2 == 1) result++;
        if ((member.DVs[StatTypes.nbAttack] / 2) % 2 == 1) result+= 2;
        if ((member.DVs[StatTypes.nbDefense] / 2) % 2 == 1) result+= 4;
        if ((member.DVs[StatTypes.nbSpeed] / 2) % 2 == 1) result+= 8;
        if ((member.DVs[StatTypes.nbSpAttack] / 2) % 2 == 1) result+= 16;
        if ((member.DVs[StatTypes.nbSpDefense] / 2) % 2 == 1) result += 32;

        return result * 40 / 63 + 30;
    }

    public int HiddenPowerStrength(StatSet DVs)
    {
        int X = 8 * ((DVs[StatTypes.nbAttack] >= 8) ? 1 : 0) + 4 * ((DVs[StatTypes.nbDefense] >= 8) ? 1 : 0) +
    2 * ((DVs[StatTypes.nbSpeed] >= 8) ? 1 : 0) + ((DVs[StatTypes.nbSpAttack] >= 8) ? 1 : 0);

        int Y = DVs[StatTypes.nbSpAttack] % 4;
        return (5 * X + Y) / 2 + 31;
    }

    public Elements HiddenPowerType(StatSet DVs)
    {
        int X = 0;
        if (DVs[StatTypes.nbHP] % 2 == 1) X += 1;
        if (DVs[StatTypes.nbAttack] % 2 == 1) X += 2;
        if (DVs[StatTypes.nbDefense] % 2 == 1) X += 4;
        if (DVs[StatTypes.nbSpeed] % 2 == 1) X += 8;
        if (DVs[StatTypes.nbSpAttack] % 2 == 1) X += 16;
        if (DVs[StatTypes.nbSpDefense] % 2 == 1) X += 32;
        X = (X*15) / 63;
        switch (X)
        {
            case 0: return Elements.nbFight; // 7
            case 1: return Elements.nbFlying; // 10
            case 2: return Elements.nbPoison; // 8
            case 3: return Elements.nbGround; // 9
            case 4: return Elements.nbRock;
            case 5: return Elements.nbBug; // 12
            case 6: return Elements.nbGhost;
            case 7: return Elements.nbSteel;
            case 8: return Elements.nbFire;
            case 9: return Elements.nbWater;
            case 10: return Elements.nbGrass;
            case 11: return Elements.nbElectr;
            case 12: return Elements.nbPsychc;
            case 13: return Elements.nbIce;
            case 14: return Elements.nbDragon;
            case 15: return Elements.nbDark;
            default: return Elements.nbNormal; // -- Should never happen, but just in case.
        }
    }
}