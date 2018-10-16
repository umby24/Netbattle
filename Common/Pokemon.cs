using System.Collections.Generic;

namespace Netbattle.Common {
    public class Pokemon {
        public int No { get; set; }
        public int GSNo { get; set; }
        public int AdvNo { get; set; }

        public bool Legendary { get; set; }
        public bool Uber { get; set; }

        public string Image { get; set; }
        public string Name { get; set; }
        public string Nickname { get; set; }

        public Elements Type1 { get; set; }
        public Elements Type2 { get; set; }

        public Traits Attribute { get; set; }
        public Traits[] PAtt { get; set; } // -- Max of 2 traits

        public byte AttNum { get; set; }
        public byte NatureNum { get; set; }
        public int Color1 { get; set; }
        public int Color2 { get; set; }
        public int[] Move { get; set; } // -- max of 4
        public byte[] MaxPP { get; set; } // -- max of 4
        public byte[] PP { get; set; } // -- max or 4

        public Items Item { get; set; }
        public Conditions Condition { get; set; }

        public int ConditionCount { get; set; }
        public byte UnownLetter { get; set; }
        public int BaseHP { get; set; }
        public int BaseAttack { get; set; }
        public int BaseDefense { get; set; }
        public int BaseSpeed { get; set; }
        public int BaseSAttack { get; set; }
        public int BaseSDefense { get; set; }
        public int BaseSpecial { get; set; }
        public int MaxHP { get; set; }
        public int HP { get; set; }
        public int Attack { get; set; }
        public int Defense { get; set; }
        public int Speed { get; set; }
        public int SpecialAttack { get; set; }
        public int SpecialDefense { get; set; }

        public byte DV_HP { get; set; }
        public byte DV_Atk { get; set; }
        public byte DV_Def { get; set; }
        public byte DV_Spd { get; set; }
        public byte DV_SAtk { get; set; }
        public byte DV_SDef { get; set; }

        public byte EV_HP { get; set; }
        public byte EV_Atk { get; set; }
        public byte EV_Def { get; set; }
        public byte EV_Spd { get; set; }
        public byte EV_SAtk { get; set; }
        public byte EV_SDef { get; set; }

        public bool Shiny { get; set; }
        public byte Level { get; set; }

        public List<Move> BaseMoves { get; set; }
        public List<Move> MachineMoves { get; set; }
        public List<Move> BreedingMoves { get; set; }
        public List<Move> RBYMoves { get; set; }
        public List<Move> RBYTM { get; set; }
        public List<Move> SpecialMoves { get; set; }

        public List<Move> AdvMoves { get; set; }
        public List<Move> ADVTM { get; set; }
        public List<Move> AdvBreeding { get; set; }
        public List<Move> AdvSpecial { get; set; }
        public List<Move> AdvTutor { get; set; }
        public List<Move> LFOnly { get; set; }
        public List<Move> MoveTutor { get; set; }
        public byte[,] MoveLevel { get; set; }

        public bool ExistRBY { get; set; }
        public bool ExistGSC { get; set; }
        public bool ExistAdv { get; set; }
        public byte StartsWith { get; set; }
        public int PercentFemale { get; set; }
        public byte Gender { get; set; }

        public int[] Evo { get; set; } // -- Max 5
        public int[] EvoM { get; set; } // -- Max 5
        public int[] Stage { get; set; } // -- Max 5

        public int MyStage { get; set; }
        public int MyMethod { get; set; }
        public int InBox { get; set; }
        public CompatModes GameVersion { get; set; }
        public int Weight { get; set; }
        public int Height { get; set; }
        public byte Offset { get; set; }
        public byte LevelBal { get; set; }
        public Items RecycleItem { get; set; }
        public byte MarkerNum { get; set; }
        public string[] Illegals { get; set; } // -- 3
        public string[] BreedIllegals { get; set; } // -- 3
        public byte EggGroup1 { get; set; }
        public byte EggGroup2 { get; set; }

        //'Rest in Slp/Frz Check
        public bool Resting { get; set; }

        //This is used for various berrys that involve randomness
        public byte ItemEffect { get; set; }

        //This is the only really strange one - it sets to their position in your lineup.
        //Used for copying info from current to team.
        public byte TeamNumber { get; set; }

        // These are all used for Database Modding
        public Traits[] ModAttr { get; set; } // -- 1
        public long TotalAdvMoves { get; set; }
        public string IllegalMod { get; set; }

        public void GetCurrentMoveset() {

        }

        public IEnumerable<Move> GetAdvMoves() {
            var result = new List<Move>();

            result.AddRange(AdvMoves);
            result.AddRange(ADVTM);
            result.AddRange(AdvBreeding);
            result.AddRange(AdvSpecial);
            result.AddRange(AdvTutor);
            result.AddRange(LFOnly);


            return result;
        }
    }
}
