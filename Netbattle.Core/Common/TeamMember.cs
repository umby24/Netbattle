using System.Text;
using System.Text.Json.Serialization;
using Netbattle.Common;
using Netbattle.Core.Database;
using Netbattle.Core.Pokemon;

namespace Netbattle.Core.Common;

public class TeamMember {
    public string Nickname { get; set; }
    public PokemonSpecies Pokemon { get; set; }
    public CompatModes GameVersion { get; set; }
    public byte UnownLetter { get; set; }
    public StatSet CurrentStats { get; set; }

    public StatSet DVs { get; set; }
    public StatSet EVs { get; set; }

    public bool Shiny { get; set; }
    public byte Level { get; set; }
    public byte Gender { get; set; }
    public int[] Move { get; set; } // -- max of 4
    public byte[] MaxPP { get; set; } // -- max of 4
    public byte[] PP { get; set; } // -- max or 4
    public byte AttNum { get; set; }
    public byte NatureNum { get; set; }

    public Items Item { get; set; }
    public Traits Attribute { get; set; }
    // -- Box attributes.
    public int InBox { get; set; }
    public byte MarkerNum { get; set; }
    
    public TeamMember(PokemonSpecies species) {
        Pokemon = species;
        CurrentStats = new StatSet();
        DVs = new StatSet();
        EVs = new StatSet();
        Move = new int[4];
        MaxPP = new byte[4];
        PP = new byte[4];
        MarkerNum = 0;
    }
    
    public IEnumerable<Move> GetAdvMoves() {
        var result = new List<Move>();

        result.AddRange(this.Pokemon.Moves.Sets[MoveTypes.nbAdvLevel]);
        result.AddRange(this.Pokemon.Moves.Sets[MoveTypes.nbAdvTutor]);
        result.AddRange(this.Pokemon.Moves.Sets[MoveTypes.nbAdvEgg]);
        result.AddRange(this.Pokemon.Moves.Sets[MoveTypes.nbAdvSpecial]);

        return result;
    }

    public void DowngradeToRBYStats() {
        DVs[StatTypes.nbHP] = 15;
        DVs[StatTypes.nbDefense] = 15;
        DVs[StatTypes.nbAttack] = 15;
        DVs[StatTypes.nbSpeed] = 15;
        DVs[StatTypes.nbSpAttack] = 15;
        DVs[StatTypes.nbSpDefense] = 15;
    }

    public void UpgradeToAdvStats() {
        DVs[StatTypes.nbHP] = 31;
        DVs[StatTypes.nbDefense] = 31;
        DVs[StatTypes.nbAttack] = 31;
        DVs[StatTypes.nbSpeed] = 31;
        DVs[StatTypes.nbSpAttack] = 31;
        DVs[StatTypes.nbSpDefense] = 31;
        NatureNum = 0;
        AttNum = 0;
        EVs[StatTypes.nbHP] = 85;
        EVs[StatTypes.nbDefense] = 85;
        EVs[StatTypes.nbAttack] = 85;
        EVs[StatTypes.nbSpeed] = 85;
        EVs[StatTypes.nbSpAttack] = 85;
        EVs[StatTypes.nbSpDefense] = 85;
    }

    public void RegenCurrentStats() {
        var ruleset = BattleSystem.GetGenerationRuleset(this.GameVersion);
        
        CurrentStats = new StatSet() {
            [StatTypes.nbHP] = ruleset.GetHp(Pokemon.BaseStats[StatTypes.nbHP], DVs[StatTypes.nbHP],
                EVs[StatTypes.nbHP], Level),
            [StatTypes.nbAttack] = ruleset.GetStat(Pokemon.BaseStats[StatTypes.nbAttack],
                DVs[StatTypes.nbAttack], EVs[StatTypes.nbAttack], Level,
                PokemonDatabase.NatureTypes[NatureNum].StatChg[0]),
            [StatTypes.nbDefense] = ruleset.GetStat(Pokemon.BaseStats[StatTypes.nbDefense],
                DVs[StatTypes.nbDefense], EVs[StatTypes.nbDefense], Level,
                PokemonDatabase.NatureTypes[NatureNum].StatChg[1]),
            [StatTypes.nbSpeed] = ruleset.GetStat(Pokemon.BaseStats[StatTypes.nbSpeed],
                DVs[StatTypes.nbSpeed], EVs[StatTypes.nbSpeed], Level,
                PokemonDatabase.NatureTypes[NatureNum].StatChg[2]),
            [StatTypes.nbSpAttack] = ruleset.GetStat(Pokemon.BaseStats[StatTypes.nbSpAttack],
                DVs[StatTypes.nbSpAttack], EVs[StatTypes.nbSpAttack], Level,
                PokemonDatabase.NatureTypes[NatureNum].StatChg[3]),
            [StatTypes.nbSpDefense] = ruleset.GetStat(Pokemon.BaseStats[StatTypes.nbSpDefense],
                DVs[StatTypes.nbSpDefense], EVs[StatTypes.nbSpDefense], Level,
                PokemonDatabase.NatureTypes[NatureNum].StatChg[4]),
        };
        
    }
    
    public void Initialize(CompatModes compatMode) {
        GameVersion = compatMode;
        if (string.IsNullOrEmpty(Nickname)) Nickname = this.Pokemon.Name;
        if (Level > 100 || Level == 0) Level = 100;
        Item = Items.nbNoItem;

        switch (Pokemon.PercentFemale) {
            case -1:
                Gender = 0;
                break;
            case 0:
                Gender = 1;
                break;
            case 16:
                Gender = 2;
                break;
        }

        if (UnownLetter > 27) UnownLetter = 0;

        switch (compatMode) {
            case CompatModes.nbTrueRuSa:
            case CompatModes.nbFullAdvance:
            case CompatModes.nbModAdv:
                if (!Item.IsAdvanceItem())
                    Item = Items.nbNoItem;

                int evTotal = EVs.Total;
                if (evTotal > 510) {
                    EVs[StatTypes.nbHP] = 85;
                    EVs[StatTypes.nbDefense] = 85;
                    EVs[StatTypes.nbAttack] = 85;
                    EVs[StatTypes.nbSpeed] = 85;
                    EVs[StatTypes.nbSpAttack] = 85;
                    EVs[StatTypes.nbSpDefense] = 85;
                }

                Pokemon.ModAttr[0] = PokemonDatabase.BasePokemonMap[Pokemon.No].ModAttr[0];
                Pokemon.ModAttr[1] = PokemonDatabase.BasePokemonMap[Pokemon.No].ModAttr[1];

                if (compatMode == CompatModes.nbModAdv) {
                    if (AttNum == 1 && Pokemon.ModAttr[1] == Traits.nbNoTrait) AttNum = 0;
                    Attribute = Pokemon.ModAttr[AttNum];
                }
                else {
                    if (AttNum == 1 && Pokemon.PAtt[1] == Traits.nbNoTrait) AttNum = 0;
                    Attribute = Pokemon.PAtt[AttNum];
                }

                if (NatureNum > 24) NatureNum = 0;
                Pokemon.MaxHP = BattleSystem.GetGenerationHp(compatMode, Pokemon.BaseStats[StatTypes.nbHP], 
                    DVs[StatTypes.nbHP], EVs[StatTypes.nbHP], Level);
                
                CurrentStats[StatTypes.nbAttack] = BattleSystem.GetGenerationStat(compatMode, Pokemon.BaseStats[StatTypes.nbAttack], DVs[StatTypes.nbAttack], EVs[StatTypes.nbAttack], Level,
                    BattleSystem.NatureStats[NatureNum].StatChg[1]);
                CurrentStats[StatTypes.nbDefense] = BattleSystem.GetGenerationStat(compatMode, Pokemon.BaseStats[StatTypes.nbDefense], DVs[StatTypes.nbDefense], EVs[StatTypes.nbDefense], Level,
                    BattleSystem.NatureStats[NatureNum].StatChg[2]);
                CurrentStats[StatTypes.nbSpeed] = BattleSystem.GetGenerationStat(compatMode, Pokemon.BaseStats[StatTypes.nbSpeed], DVs[StatTypes.nbSpeed], EVs[StatTypes.nbSpeed], Level,
                    BattleSystem.NatureStats[NatureNum].StatChg[3]);
                CurrentStats[StatTypes.nbSpAttack] = BattleSystem.GetGenerationStat(compatMode, Pokemon.BaseStats[StatTypes.nbSpAttack], DVs[StatTypes.nbSpAttack], EVs[StatTypes.nbSpAttack], Level,
                    BattleSystem.NatureStats[NatureNum].StatChg[4]);
                CurrentStats[StatTypes.nbSpDefense] = BattleSystem.GetGenerationStat(compatMode, Pokemon.BaseStats[StatTypes.nbSpDefense], DVs[StatTypes.nbSpDefense], EVs[StatTypes.nbSpDefense], Level,
                    BattleSystem.NatureStats[NatureNum].StatChg[5]);

                break;
            default:
                DowngradeToRBYStats();
                CurrentStats[StatTypes.nbAttack] = BattleSystem.GetGenerationStat(compatMode, Pokemon.BaseStats[StatTypes.nbAttack], DVs[StatTypes.nbAttack], EVs[StatTypes.nbAttack], Level,
                    BattleSystem.NatureStats[NatureNum].StatChg[1]);
                CurrentStats[StatTypes.nbDefense] = BattleSystem.GetGenerationStat(compatMode, Pokemon.BaseStats[StatTypes.nbDefense], DVs[StatTypes.nbDefense], EVs[StatTypes.nbDefense], Level,
                    BattleSystem.NatureStats[NatureNum].StatChg[2]);
                CurrentStats[StatTypes.nbSpeed] = BattleSystem.GetGenerationStat(compatMode, Pokemon.BaseStats[StatTypes.nbSpeed], DVs[StatTypes.nbSpeed], EVs[StatTypes.nbSpeed], Level,
                    BattleSystem.NatureStats[NatureNum].StatChg[3]);
                CurrentStats[StatTypes.nbSpAttack] = BattleSystem.GetGenerationStat(compatMode, Pokemon.BaseStats[StatTypes.nbSpAttack], DVs[StatTypes.nbSpAttack], EVs[StatTypes.nbSpAttack], Level,
                    BattleSystem.NatureStats[NatureNum].StatChg[4]);
                CurrentStats[StatTypes.nbSpDefense] = BattleSystem.GetGenerationStat(compatMode, Pokemon.BaseStats[StatTypes.nbSpDefense], DVs[StatTypes.nbSpDefense], EVs[StatTypes.nbSpDefense], Level,
                    BattleSystem.NatureStats[NatureNum].StatChg[5]);
                
                if (GameVersion == 0 || GameVersion == (CompatModes)5) {
                    Item = Items.nbNoItem;
                    Shiny = false;
                }
                else {
                    Shiny = BattleSystem.ShinyDV((short)DVs[StatTypes.nbHP]) && DVs[StatTypes.nbDefense] == 10 && DVs[StatTypes.nbSpeed] == 10 && DVs[StatTypes.nbSpAttack] == 10;
                }

                DVs[StatTypes.nbHP] = 0;
                if (DVs[StatTypes.nbAttack] % 2 == 1) DVs[StatTypes.nbHP] += 8;
                if (DVs[StatTypes.nbDefense] % 2 == 1) DVs[StatTypes.nbHP] += 4;
                if (DVs[StatTypes.nbSpeed] % 2 == 1) DVs[StatTypes.nbHP] += 2;
                if (DVs[StatTypes.nbSpAttack] % 2 == 1) DVs[StatTypes.nbHP] += 1;
                Pokemon.MaxHP = BattleSystem.GetGenerationHp(compatMode, Pokemon.BaseStats[StatTypes.nbHP], 
                    DVs[StatTypes.nbHP], EVs[StatTypes.nbHP], Level);
                CurrentStats[StatTypes.nbHP] = Pokemon.MaxHP;
                
                if (Pokemon.PercentFemale == -1) {
                    Gender = 0;
                }
                else {
                    if (DVs[StatTypes.nbAttack] <= Pokemon.PercentFemale - 1) {
                        Gender = 2;
                    }
                    else {
                        Gender = 1;
                    }
                }

                Attribute = Traits.nbNoTrait;
                if (Item > (Items)41) Item = Items.nbNoItem;
                if (Pokemon.No == 201) {
                    UnownLetter =
                        (byte)((DVs[StatTypes.nbAttack] % 8 / 2 * 64 + DVs[StatTypes.nbDefense] % 8 / 2 * 16 + (DVs[StatTypes.nbSpeed] % 8) / 2 * 4 + (DVs[StatTypes.nbSpAttack] % 8) / 2) /
                            10 % 26);
                }
                else {
                    UnownLetter = 0;
                }

                break;
        }
    }

    public static TeamMember FromBinary(string pokeData, string nickname) {
        var No = NbMethods.Bin2Dec(pokeData.Substring(0, 9));

        var resultPokemon = new TeamMember(PokemonDatabase.BasePokemon[No - 1]);
        resultPokemon.Nickname = nickname;

        var GvBits = NbMethods.Bin2Dec(pokeData.Substring(9, 3));
        resultPokemon.GameVersion = (CompatModes)GvBits;

        var LvBits = NbMethods.Bin2Dec(pokeData.Substring(12, 7));
        resultPokemon.Level = (byte)LvBits;

        var itemBits = NbMethods.Bin2Dec(pokeData.Substring(19, 7));
        resultPokemon.Item = (Items)itemBits;

        var natureBits = NbMethods.Bin2Dec(pokeData.Substring(26, 5));
        resultPokemon.NatureNum = (byte)natureBits;
        var AttBit = NbMethods.Bin2Dec(pokeData.Substring(31, 1));
        resultPokemon.AttNum = (byte)AttBit;
        var GenderBit = NbMethods.Bin2Dec(pokeData.Substring(32, 1));
        resultPokemon.Gender = (byte)(GenderBit + 1); // -- Somehow missed the +1 in the original code. :)
        var ShinyBit = NbMethods.Bin2Dec(pokeData.Substring(33, 1));
        resultPokemon.Shiny = ShinyBit > 0;
        var InboxBits = NbMethods.Bin2Dec(pokeData.Substring(34, 4));
        resultPokemon.InBox = InboxBits;
        var unownBits = NbMethods.Bin2Dec(pokeData.Substring(38, 5));
        resultPokemon.UnownLetter = (byte)unownBits;
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
        resultPokemon.DVs = new StatSet() {
            [StatTypes.nbHP] = dvhp,
            [StatTypes.nbAttack] = dvatk,
            [StatTypes.nbDefense] = dvdef,
            [StatTypes.nbSpeed] = dvspd,
            [StatTypes.nbSpAttack] = dvsatk,
            [StatTypes.nbSpDefense] = dvsdef
        };

        var evhp = NbMethods.Bin2Dec(pokeData.Substring(109, 8));
        var evatk = NbMethods.Bin2Dec(pokeData.Substring(117, 8));
        var evdef = NbMethods.Bin2Dec(pokeData.Substring(125, 8));
        var evspd = NbMethods.Bin2Dec(pokeData.Substring(133, 8));
        var evsatk = NbMethods.Bin2Dec(pokeData.Substring(141, 8));
        var evsdef = NbMethods.Bin2Dec(pokeData.Substring(149, 8));
        resultPokemon.EVs = new StatSet() {
            [StatTypes.nbHP] = evhp,
            [StatTypes.nbAttack] = evatk,
            [StatTypes.nbDefense] = evdef,
            [StatTypes.nbSpeed] = evspd,
            [StatTypes.nbSpAttack] = evsatk,
            [StatTypes.nbSpDefense] = evsdef
        };

        return resultPokemon;
    }

    public byte[] ToStringBytes() {
        // -- Code.bas - 6078 AKA PKMN2Str
        var sb = new StringBuilder();

        sb.Append(NbMethods.Dec2Bin(Pokemon.No, 9));
        sb.Append(NbMethods.Dec2Bin((int)GameVersion, 3));
        sb.Append(NbMethods.Dec2Bin(Level, 7));
        sb.Append(NbMethods.Dec2Bin((int)Item, 7));
        sb.Append(NbMethods.Dec2Bin(NatureNum, 5));
        sb.Append(AttNum.ToString());
        sb.Append((Gender - 1 == -1) ? 0 : Gender - 1);
        sb.Append(Shiny ? 1 : 0);
        sb.Append(NbMethods.Dec2Bin(InBox, 4));
        sb.Append(NbMethods.Dec2Bin(UnownLetter, 5));
        sb.Append(NbMethods.Dec2Bin(Move[0], 9));
        sb.Append(NbMethods.Dec2Bin(Move[1], 9));
        sb.Append(NbMethods.Dec2Bin(Move[2], 9));
        sb.Append(NbMethods.Dec2Bin(Move[3], 9));
        sb.Append(NbMethods.Dec2Bin(DVs[StatTypes.nbHP], 5));
        sb.Append(NbMethods.Dec2Bin(DVs[StatTypes.nbAttack], 5));
        sb.Append(NbMethods.Dec2Bin(DVs[StatTypes.nbDefense], 5));
        sb.Append(NbMethods.Dec2Bin(DVs[StatTypes.nbSpeed], 5));
        sb.Append(NbMethods.Dec2Bin(DVs[StatTypes.nbSpAttack], 5));
        sb.Append(NbMethods.Dec2Bin(DVs[StatTypes.nbSpDefense], 5));
        // -- EVs..
        sb.Append(NbMethods.Dec2Bin(EVs[StatTypes.nbHP], 8));
        sb.Append(NbMethods.Dec2Bin(EVs[StatTypes.nbAttack], 8));
        sb.Append(NbMethods.Dec2Bin(EVs[StatTypes.nbDefense], 8));
        sb.Append(NbMethods.Dec2Bin(EVs[StatTypes.nbSpeed], 8));
        sb.Append(NbMethods.Dec2Bin(EVs[StatTypes.nbSpAttack], 8));
        sb.Append(NbMethods.Dec2Bin(EVs[StatTypes.nbSpDefense], 8));

        byte[] nameBytes = Encoding.ASCII.GetBytes(Nickname.PadRight(15));
        byte[] pokeBytes = NbMethods.BinaryToBytes(sb.ToString());

        var final = new byte[15 + pokeBytes.Length];
        Buffer.BlockCopy(nameBytes, 0, final, 0, 15);
        Buffer.BlockCopy(pokeBytes, 0, final, 15, pokeBytes.Length);

        return final;
    }
}