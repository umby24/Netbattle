using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;

namespace Netbattle.Common {
    public class Pokemon {
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

        public string Nickname { get; set; }
        [JsonIgnore]
        public Elements Type1 { get; set; }
        [JsonIgnore]
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
        [JsonIgnore]
        public int BaseHP { get; set; }
        [JsonIgnore]
        public int BaseAttack { get; set; }
        [JsonIgnore]
        public int BaseDefense { get; set; }
        [JsonIgnore]
        public int BaseSpeed { get; set; }
        [JsonIgnore]
        public int BaseSAttack { get; set; }
        [JsonIgnore]
        public int BaseSDefense { get; set; }
        [JsonIgnore]
        public int BaseSpecial { get; set; }
        [JsonIgnore]
        public int MaxHP { get; set; }
        [JsonIgnore]
        public int HP { get; set; }
        [JsonIgnore]
        public int Attack { get; set; }
        [JsonIgnore]
        public int Defense { get; set; }
        [JsonIgnore]
        public int Speed { get; set; }
        [JsonIgnore]
        public int SpecialAttack { get; set; }
        [JsonIgnore]
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
        [JsonIgnore]
        public List<Move> BaseMoves { get; set; }
        [JsonIgnore]
        public List<Move> MachineMoves { get; set; }
        [JsonIgnore]
        public List<Move> BreedingMoves { get; set; }
        [JsonIgnore]
        public List<Move> RBYMoves { get; set; }
        [JsonIgnore]
        public List<Move> RBYTM { get; set; }
        [JsonIgnore]
        public List<Move> SpecialMoves { get; set; }
        [JsonIgnore]
        public List<Move> AdvMoves { get; set; }
        [JsonIgnore]
        public List<Move> ADVTM { get; set; }
        [JsonIgnore]
        public List<Move> AdvBreeding { get; set; }
        [JsonIgnore]
        public List<Move> AdvSpecial { get; set; }
        [JsonIgnore]
        public List<Move> AdvTutor { get; set; }
        [JsonIgnore]
        public List<Move> LFOnly { get; set; }
        [JsonIgnore]
        public List<Move> MoveTutor { get; set; }
        [JsonIgnore]
        public byte[,] MoveLevel { get; set; }
        [JsonIgnore]
        public bool ExistRBY { get; set; }
        [JsonIgnore]
        public bool ExistGSC { get; set; }
        [JsonIgnore]
        public bool ExistAdv { get; set; }
        [JsonIgnore]
        public byte StartsWith { get; set; }
        [JsonIgnore]
        public int PercentFemale { get; set; }

        public byte Gender { get; set; }
        [JsonIgnore]
        public int[] Evo { get; set; } // -- Max 5
        [JsonIgnore]
        public int[] EvoM { get; set; } // -- Max 5
        [JsonIgnore]
        public int[] Stage { get; set; } // -- Max 5

        public int MyStage { get; set; }
        public int MyMethod { get; set; }
        public int InBox { get; set; }
        public CompatModes GameVersion { get; set; }
        [JsonIgnore]
        public int Weight { get; set; }
        [JsonIgnore]
        public int Height { get; set; }
        [JsonIgnore]
        public byte Offset { get; set; }
        [JsonIgnore]
        public byte LevelBal { get; set; }
        public Items RecycleItem { get; set; }
        public byte MarkerNum { get; set; }
        public string[] Illegals { get; set; } // -- 3
        public string[] BreedIllegals { get; set; } // -- 3
        public byte EggGroup1 { get; set; }
        public byte EggGroup2 { get; set; }

        //'Rest in Slp/Frz Check
        [JsonIgnore]
        public bool Resting { get; set; }

        //This is used for various berrys that involve randomness
        [JsonIgnore]
        public byte ItemEffect { get; set; }

        //This is the only really strange one - it sets to their position in your lineup.
        //Used for copying info from current to team.
        [JsonIgnore]
        public byte TeamNumber { get; set; }

        // These are all used for Database Modding
        [JsonIgnore]
        public Traits[] ModAttr { get; set; } // -- 1
        [JsonIgnore]
        public long TotalAdvMoves { get; set; }
        [JsonIgnore]
        public string IllegalMod { get; set; }

        public Pokemon() {
            Move = new int[4];

        }

        /// <summary>
        /// This method will take the newest pokemon from the database, and replay any user-choice fields on top of that object, and return it.
        /// </summary>
        /// <returns></returns>
        public Pokemon SetupFromDatabase() {
            Pokemon newPoke = PokemonDatabase.BasePokemon[No - 1];
            newPoke.Gender = Gender;
            newPoke.Level = Level;
            newPoke.Shiny = Shiny;
            newPoke.UnownLetter = UnownLetter;
            newPoke.Condition = Condition;
            newPoke.Item = Item;
            newPoke.Nickname = Nickname;
            newPoke.Attribute = Attribute;
            newPoke.Move = Move;
            newPoke.MaxPP = MaxPP;
            newPoke.PP = PP;
            newPoke.NatureNum = NatureNum;
            newPoke.AttNum = AttNum;
            newPoke.EV_Atk = EV_Atk;
            newPoke.EV_Def = EV_Def;
            newPoke.EV_HP = EV_HP;
            newPoke.EV_SAtk = EV_SAtk;
            newPoke.EV_SDef = EV_SDef;
            newPoke.EV_Spd = EV_Spd;
            newPoke.DV_Atk = DV_Atk;
            newPoke.DV_Def = DV_Def;
            newPoke.DV_HP = DV_HP;
            newPoke.DV_SAtk = DV_SAtk;
            newPoke.DV_SDef = DV_SDef;
            newPoke.DV_Spd = DV_Spd;

            return newPoke;
        }

        public void GetCurrentMoveset() {

        }

        public static Pokemon FromBinary(string pokeData, string nickname) {
            
            var No = NbMethods.Bin2Dec(pokeData.Substring(0, 9));

            var resultPokemon = PokemonDatabase.BasePokemon[No - 1];
            resultPokemon.Nickname = nickname;

            var GvBits = NbMethods.Bin2Dec(pokeData.Substring(9, 3));
            resultPokemon.GameVersion = (CompatModes)GvBits;

            var LvBits = NbMethods.Bin2Dec(pokeData.Substring(12, 7));
            resultPokemon.Level = (byte) LvBits;

            var itemBits = NbMethods.Bin2Dec(pokeData.Substring(19, 7));
            resultPokemon.Item = (Items) itemBits;

            var natureBits = NbMethods.Bin2Dec(pokeData.Substring(26, 5));
            resultPokemon.NatureNum = (byte) natureBits;
            var AttBit = NbMethods.Bin2Dec(pokeData.Substring(31, 1));
            resultPokemon.AttNum = (byte) AttBit;
            var GenderBit = NbMethods.Bin2Dec(pokeData.Substring(32, 1));
            resultPokemon.Gender = (byte)GenderBit;
            var ShinyBit = NbMethods.Bin2Dec(pokeData.Substring(33, 1));
            resultPokemon.Shiny = ShinyBit > 0;
            var InboxBits = NbMethods.Bin2Dec(pokeData.Substring(34, 4));
            resultPokemon.InBox = InboxBits;
            var unownBits = NbMethods.Bin2Dec(pokeData.Substring(38, 5));
            resultPokemon.UnownLetter = (byte) unownBits;
            var move0 = NbMethods.Bin2Dec(pokeData.Substring(43, 9));
            resultPokemon.Move[0] = move0;
            var move1 = NbMethods.Bin2Dec(pokeData.Substring(52, 9));
            resultPokemon.Move[1] = move1;
            var move2 = NbMethods.Bin2Dec(pokeData.Substring(61, 9));
            resultPokemon.Move[2] = move2;
            var move3 = NbMethods.Bin2Dec(pokeData.Substring(70, 9));
            resultPokemon.Move[3] = move3;

            var dvhp = NbMethods.Bin2Dec(pokeData.Substring(79, 5));
            var dvatk = NbMethods.Bin2Dec(pokeData.Substring(84, 5));
            var dvdef = NbMethods.Bin2Dec(pokeData.Substring(89, 5));
            var dvspd = NbMethods.Bin2Dec(pokeData.Substring(94, 5));
            var dvsatk = NbMethods.Bin2Dec(pokeData.Substring(99, 5));
            var dvsdef = NbMethods.Bin2Dec(pokeData.Substring(104, 5));
            resultPokemon.DV_HP = (byte) dvhp;
            resultPokemon.DV_Atk = (byte)dvatk;
            resultPokemon.DV_Def = (byte) dvdef;
            resultPokemon.DV_Spd = (byte) dvspd;
            resultPokemon.DV_SAtk = (byte) dvsatk;
            resultPokemon.DV_SDef = (byte) dvsdef;

            var evhp = NbMethods.Bin2Dec(pokeData.Substring(109, 8));
            var evatk = NbMethods.Bin2Dec(pokeData.Substring(117, 8));
            var evdef = NbMethods.Bin2Dec(pokeData.Substring(125, 8));
            var evspd = NbMethods.Bin2Dec(pokeData.Substring(133, 8));
            var evsatk = NbMethods.Bin2Dec(pokeData.Substring(141, 8));
            var evsdef = NbMethods.Bin2Dec(pokeData.Substring(149, 8));
            resultPokemon.EV_HP = (byte) evhp;
            resultPokemon.EV_Atk = (byte) evatk;
            resultPokemon.EV_Def = (byte) evdef;
            resultPokemon.EV_Spd = (byte) evspd;
            resultPokemon.EV_SAtk = (byte) evsatk;
            resultPokemon.EV_SDef = (byte) evsdef;

            return resultPokemon;
        }

        public byte[] ToStringBytes() {// -- Code.bas - 6078
            var sb = new StringBuilder();
            
            sb.Append(NbMethods.Dec2Bin(No, 9));
            sb.Append(NbMethods.Dec2Bin((int) GameVersion, 3));
            sb.Append(NbMethods.Dec2Bin(Level, 7));
            sb.Append(NbMethods.Dec2Bin((int)Item, 7));
            sb.Append(NbMethods.Dec2Bin(NatureNum, 5));
            sb.Append(AttNum.ToString());
            sb.Append(Gender);
            sb.Append(Shiny ? 1 : 0);
            sb.Append(NbMethods.Dec2Bin(InBox, 4));
            sb.Append(NbMethods.Dec2Bin(UnownLetter, 5));
            sb.Append(NbMethods.Dec2Bin(Move[0], 9));
            sb.Append(NbMethods.Dec2Bin(Move[1], 9));
            sb.Append(NbMethods.Dec2Bin(Move[2], 9));
            sb.Append(NbMethods.Dec2Bin(Move[3], 9));
            sb.Append(NbMethods.Dec2Bin(DV_HP, 5));
            sb.Append(NbMethods.Dec2Bin(DV_Atk, 5));
            sb.Append(NbMethods.Dec2Bin(DV_Def, 5));
            sb.Append(NbMethods.Dec2Bin(DV_Spd, 5));
            sb.Append(NbMethods.Dec2Bin(DV_SAtk, 5));
            sb.Append(NbMethods.Dec2Bin(DV_SDef, 5));
            // -- EVs..
            sb.Append(NbMethods.Dec2Bin(EV_HP, 8));
            sb.Append(NbMethods.Dec2Bin(EV_Atk, 8));
            sb.Append(NbMethods.Dec2Bin(EV_Def, 8));
            sb.Append(NbMethods.Dec2Bin(EV_Spd, 8));
            sb.Append(NbMethods.Dec2Bin(EV_SAtk, 8));
            sb.Append(NbMethods.Dec2Bin(EV_SDef, 8));

            byte[] nameBytes = Encoding.ASCII.GetBytes(Nickname.PadRight(15));
            byte[] pokeBytes = NbMethods.BinaryToBytes(sb.ToString());

            var final = new byte[15 + pokeBytes.Length];
            Buffer.BlockCopy(nameBytes, 0, final, 0, 15);
            Buffer.BlockCopy(pokeBytes, 0, final, 15, pokeBytes.Length);

            return final;
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
