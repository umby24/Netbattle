using Netbattle.Core.Common;
using Netbattle.Core.Database;

namespace Netbattle.Core.Battle;

// -- Only diff here is "FR/LG only rules".
public class ADVPlusRules : ADVRules 
{
    public new bool CheckMove(PokemonSpecies species, Move move)
    {
        var result = species.Moves.Sets[MoveTypes.nbAdvLevel].Contains(move);
        
        if ( species.Moves.Sets[MoveTypes.nbAdvTM].Contains(move)) result = true;
        if ( species.Moves.Sets[MoveTypes.nbAdvEgg].Contains(move)) result = true;
        if ( species.Moves.Sets[MoveTypes.nbAdvSpecial].Contains(move)) result = true;
        if ( species.Moves.Sets[MoveTypes.nbAdvTutor].Contains(move)) result = true;
        if (species.Moves.Sets[MoveTypes.nbAdvFL].Contains(move)) result = true;
        
        return result;
    }
}