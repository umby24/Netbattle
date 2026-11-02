using System.Text.Json.Serialization;
using Netbattle.Core.Database;
using Netbattle.Core.Pokemon;

namespace Netbattle.Core.Common;

public class PokemonSpecies
{
    public int No { get; set; }
    [JsonIgnore]
    public int GSNo { get; set; }
    [JsonIgnore]
    public int AdvNo { get; set; }
    [JsonIgnore]
    public bool Legendary { get; set; }
    [JsonIgnore]
    public bool Uber { get; set; }
    [JsonIgnore]
    public string Image { get; set; }
    [JsonIgnore]
    public string Name { get; set; }
    [JsonIgnore]
    public Elements Type1 { get; set; }
    [JsonIgnore]
    public Elements Type2 { get; set; }
    
    public Traits[] PAtt { get; set; } // -- Max of 2 traits


    public int Color1 { get; set; }
    public int Color2 { get; set; }
    
    public StatSet BaseStats { get; set; }
    [JsonIgnore]
    public int BaseSpecial { get; set; }
    [JsonIgnore]
    public int MaxHP { get; set; }

    public LearnSet Moves { get; set; }
    
    [JsonIgnore]
    public byte[,] MoveLevel { get; set; }
    [JsonIgnore]
    public bool ExistRBY { get; set; }
    [JsonIgnore]
    public bool ExistGSC { get; set; }

    public bool ExistADV { get; set; }

    [JsonIgnore]
    public int Weight { get; set; }
    [JsonIgnore]
    public int Height { get; set; }
    [JsonIgnore]
    public int PercentFemale { get; set; }
    [JsonIgnore]
    public int[] Evo { get; set; } // -- Max 5
    [JsonIgnore]
    public int[] EvoM { get; set; } // -- Max 5
    [JsonIgnore]
    public int[] Stage { get; set; } // -- Max 5

    public int MyStage { get; set; }
    
    // These are all used for Database Modding
    [JsonIgnore]
    public Traits[] ModAttr { get; set; } // -- 1
    [JsonIgnore]
    public string IllegalMod { get; set; }
    
    public int MyMethod { get; set; } // -- How is this pokemon evolved? 0 = none,
                
    public string[] Illegals { get; set; } // -- 3
    public string[] BreedIllegals { get; set; } // -- 3
    public byte EggGroup1 { get; set; }
    public byte EggGroup2 { get; set; }
    
     public List<Move> GetCurrentMoveset(CompatModes compatMode) {
            List<Move> legalMoveSet = new List<Move>();
            
            switch (compatMode)
            {
                case CompatModes.nbTrueRBY:
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbRBYLevel]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbRBYTM]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbGSCSpecial]);
                    break;
                case CompatModes.nbTrueGSC:
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbGSCLevel]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbGSCTM]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbGSCEgg]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbGSCTutor]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbGSCSpecial]);
                    break;
                case CompatModes.nbRBYTrade:
                case CompatModes.nbGSCTrade:
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbRBYLevel]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbRBYTM]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbGSCSpecial]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbGSCTM]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbGSCEgg]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbGSCTutor]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbGSCSpecial]);
                    break;
                case CompatModes.nbTrueRuSa:
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbAdvLevel]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbAdvTM]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbAdvSpecial]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbAdvEgg]);
                    break;
                case CompatModes.nbFullAdvance:
                case CompatModes.nbModAdv:
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbAdvLevel]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbAdvTM]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbAdvSpecial]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbAdvEgg]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbAdvFL]);
                    legalMoveSet.AddRange(this.Moves.Sets[MoveTypes.nbAdvTutor]);
                    break;
            }
            
            return legalMoveSet;
        }
}