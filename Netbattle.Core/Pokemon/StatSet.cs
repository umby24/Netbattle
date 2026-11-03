namespace Netbattle.Core.Pokemon;
using Netbattle.Core.Common;
using Newtonsoft.Json;


public class StatSet
{
    // -- Named properties are what get written to team.jnb, the indexer is invisible to the serializer.
    public int HP { get; set; }
    public int Attack { get; set; }
    public int Defense { get; set; }
    public int Speed { get; set; }
    public int SpAttack { get; set; }
    public int SpDefense { get; set; }

    public int this[Netbattle.Core.Common.StatTypes stat]
    {
        get => stat switch
        {
            StatTypes.nbHP => HP,
            StatTypes.nbAttack => Attack,
            StatTypes.nbDefense => Defense,
            StatTypes.nbSpeed => Speed,
            StatTypes.nbSpAttack => SpAttack,
            StatTypes.nbSpDefense => SpDefense,
            _ => throw new ArgumentOutOfRangeException(nameof(stat), stat, null)
        };
        set
        {
            switch (stat)
            {
                case StatTypes.nbHP: HP = value; break;
                case StatTypes.nbAttack: Attack = value; break;
                case StatTypes.nbDefense: Defense = value; break;
                case StatTypes.nbSpeed: Speed = value; break;
                case StatTypes.nbSpAttack: SpAttack = value; break;
                case StatTypes.nbSpDefense: SpDefense = value; break;
                default: throw new ArgumentOutOfRangeException(nameof(stat), stat, null);
            }
        }
    }

    public static readonly StatTypes[] All = Enum.GetValues(typeof(StatTypes)).Cast<StatTypes>().ToArray();

    [JsonIgnore]
    public int Total
    {
        get { int sum = 0; foreach(var s in All) sum += this[s]; return sum; }
    }
}
