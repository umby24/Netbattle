using System.Text.Json.Serialization;

namespace Netbattle.Core.Common;

public class BattlePokemon
{
    public Conditions Condition { get; set; }

    public int ConditionCount { get; set; }
    //'Rest in Slp/Frz Check
    [JsonIgnore]
    public bool Resting { get; set; }

    //This is used for various berrys that involve randomness
    [JsonIgnore]
    public byte ItemEffect { get; set; }
}