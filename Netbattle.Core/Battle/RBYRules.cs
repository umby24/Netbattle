using Netbattle.Core.Common;
using Netbattle.Core.Database;
using Netbattle.Core.Pokemon;

namespace Netbattle.Core.Battle;

public class RBYRules : IGenerationRuleset
{
    public static readonly int DV_MAX = 15;
    
    public int GetStat(int baseVal, int iv, int ev, int level, int pmod) {
        // -- EV, Nature modifier unused. 'iv' is 'dv'.
        return (int)(level * (baseVal + iv + 31.9f) / 50) + 5;
    }

    public int GetHp(int baseVal, int iv, int ev, int level) {
        return (int)(level * (baseVal + iv + 31.9f + 50) / 50) + 10;
    }

    public int HiddenPowerStrength(StatSet DVs) {
        int X = 8 * ((DVs[StatTypes.nbAttack] >= 8) ? 1 : 0) + 4 * ((DVs[StatTypes.nbDefense] >= 8) ? 1 : 0) + 
                2 * ((DVs[StatTypes.nbSpeed] >= 8) ? 1 : 0) + ((DVs[StatTypes.nbSpAttack] >= 8) ? 1 : 0);

        int Y = DVs[StatTypes.nbSpAttack] % 4;
        return (5* X + Y) / 2 + 31;
    }

    public Elements HiddenPowerType(StatSet DVs) {
        int index = 4 * (DVs[StatTypes.nbAttack] % 4) + (DVs[StatTypes.nbDefense] % 4);
        switch (index)
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

    public bool CheckMove(PokemonSpecies species, Move move)
    {
        var result = species.Moves.Sets[MoveTypes.nbRBYLevel].Contains(move) 
                     || species.Moves.Sets[MoveTypes.nbRBYTM].Contains(move) 
                     || species.No is 54 or 55 && move.ID == 6;
        
        // -- Amnesia Psyduck exception
        return result;
    }
}