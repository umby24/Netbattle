namespace Netbattle.Core.Pokemon;
using Netbattle.Core.Common;


public class StatSet
{
    private int _hp, _atk, _def, _spd, _spAtk, _spDef;

    public int this[Netbattle.Core.Common.StatTypes stat]
    {
        get => stat switch
        {
            StatTypes.nbHP => _hp,
            StatTypes.nbAttack => _atk,
            StatTypes.nbDefense => _def,
            StatTypes.nbSpeed => _spd,
            StatTypes.nbSpAttack => _spAtk,
            StatTypes.nbSpDefense => _spDef,
            _ => throw new ArgumentOutOfRangeException(nameof(stat), stat, null)
        };
        set
        {
            switch (stat)
            {
                case StatTypes.nbHP: _hp = value; break;
                case StatTypes.nbAttack: _atk = value; break;
                case StatTypes.nbDefense: _def = value; break;
                case StatTypes.nbSpeed: _spd = value; break;
                case StatTypes.nbSpAttack: _spAtk = value; break;
                case StatTypes.nbSpDefense: _spDef = value; break;
                default: throw new ArgumentOutOfRangeException(nameof(stat), stat, null);
            }
        }
    }
    
    public static readonly StatTypes[] All = Enum.GetValues(typeof(StatTypes)).Cast<StatTypes>().ToArray();

    public int Total
    {
        get { int sum = 0; foreach(var s in All) sum += this[s]; return sum; }
    }
}