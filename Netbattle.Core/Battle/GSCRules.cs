using Netbattle.Core.Common;
using Netbattle.Core.Database;
using Netbattle.Core.Pokemon;

namespace Netbattle.Core.Battle;

public class GSCRules : IGenerationRuleset
{
    public int GetStat(int baseVal, int iv, int ev, int level, int pmod) {
        // -- EV, Nature modifier unused. 'iv' is 'dv'.
        return (int)(level * (baseVal + iv + 31.9f) / 50) + 5;
    }

    public int GetHp(int baseVal, int iv, int ev, int level) {
        return (int)(level * (baseVal + iv + 31.9f + 50) / 50) + 10;
    }

    public bool CheckMove(PokemonSpecies species, Move move)
    {
        
        var result = species.Moves.Sets[MoveTypes.nbGSCLevel].Contains(move);
        if (species.Moves.Sets[MoveTypes.nbGSCEgg].Contains(move)) result = true;
        if (species.Moves.Sets[MoveTypes.nbGSCTM].Contains(move)) result = true;
        if (species.Moves.Sets[MoveTypes.nbGSCSpecial].Contains(move)) result = true;
        if (species.Moves.Sets[MoveTypes.nbGSCTutor].Contains(move)) result = true;
            
        switch (species.No)
        {
            case 83 when move.ID == 12:
            case 207 when move.ID == 55:
                result = true; // -- Farfetch'd can learn Cut in GSC.
                break;
        }
        // -- In GSC, ExtremeSpeed and Twister can be learned by different pokes.
        // -- In the original code, there was a bug here!! that allowed Moltres to learn these instead of Dragonair.
        switch (move.ID)
        {
            case 61: // -- ExtremeSpeed
            case 240: // -- Twister
                return species.No is 147 or 148 or 149;
            case 123 when species.No == 83: // -- Gligar Earthquake
                return true;
            case 55 when species.No == 207: // -- Farfetch'd Cut
                return true;
            case 45: // -- Dizzy Punch
                return species.No is 239 or 125
                                         // -- Pichu, Pikachu, Raichu
                                         or 172 or 25 or 26
                                         // -- Cleffa, Clefariy, Clefable
                                         or 173 or 35 or 36
                                                      // -- Igglybuff, Jigglypuff, Wigglytuff
                or 174 or 39 or 40
                                                      // -- Smoochum, Jynx
                or 238 or 124
                       // -- Magby, Magmar
                or 240 or 126
                       // -- Tyrogue, Hitmonlee, Hitmonchan, Hitmontop
                or 236 or 106 or 107 or 237;
            default:
                return result;
        }
    }

    public int HiddenPowerStrength(StatSet DVs)
    {
        int X = 8 * ((DVs[StatTypes.nbAttack] >= 8) ? 1 : 0) + 4 * ((DVs[StatTypes.nbDefense] >= 8) ? 1 : 0) + 
            2 * ((DVs[StatTypes.nbSpeed] >= 8) ? 1 : 0) + ((DVs[StatTypes.nbSpAttack] >= 8) ? 1 : 0);

        int Y = DVs[StatTypes.nbSpAttack] % 4;
        return (5* X + Y) / 2 + 31;
    }

    public Elements HiddenPowerType(StatSet DVs)
    {
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
}