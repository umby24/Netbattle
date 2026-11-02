using Netbattle.Core.Common;
using Netbattle.Core.Database;

namespace Netbattle.Core.Pokemon;

public class LearnSet
{
    public Dictionary<MoveTypes, IReadOnlyList<Move>> Sets { get; set; }
}