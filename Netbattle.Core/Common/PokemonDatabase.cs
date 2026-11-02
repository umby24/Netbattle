using Netbattle.Core.Common;
using Netbattle.Core.Database;
using Netbattle.Core.Files;
using Netbattle.Core.Pokemon;

namespace Netbattle.Common {
    public class PokemonDatabase {
        public static List<PokemonSpecies> BasePokemon = new List<PokemonSpecies>();
        public static Dictionary<int, PokemonSpecies> BasePokemonMap = new Dictionary<int, PokemonSpecies>();
        public static Dictionary<int, PokedexInfo> Pokedex = new Dictionary<int, PokedexInfo>();
        public static List<NatureType> NatureTypes = new List<NatureType>();
        public PokemonSpecies this[int id] => BasePokemonMap[id];
        public PokemonSpecies this[string id] => BasePokemonMap.FirstOrDefault(a => string.Equals(a.Value.Name, id, StringComparison.OrdinalIgnoreCase)).Value;
        
        public static void Load() {
            var pokeDb = new CdbFile("PokeDB.cdb");
            pokeDb.Load();
            Logger.Log(LogType.Verbose, "PokemonDB Read and decompressed.");

            foreach (string[] entry in pokeDb.LineContent) {
                ParsePokedexInfo(entry);
                ParsePokemonDbLine(entry);
            }
            BuildMap();
            LoadNatures();
            Logger.Log(LogType.Info, $"Pokemon Database loaded successfully. Showing {BasePokemon.Count} Pokemon.");
        }

        private static void LoadNatures()
        {
            NatureTypes.Add(new NatureType("Hardy"));
            NatureTypes.Add(new NatureType("Lonely"));
            NatureTypes.Add(new NatureType("Brave"));
            NatureTypes.Add(new NatureType("Adamant"));
            NatureTypes.Add(new NatureType("Naughty"));
            NatureTypes.Add(new NatureType("Bold"));
            NatureTypes.Add(new NatureType("Docile"));
            NatureTypes.Add(new NatureType("Relaxed"));
            NatureTypes.Add(new NatureType("Impish"));
            NatureTypes.Add(new NatureType("Lax"));
            NatureTypes.Add(new NatureType("Timid"));
            NatureTypes.Add(new NatureType("Hasty"));
            NatureTypes.Add(new NatureType("Serious"));
            NatureTypes.Add(new NatureType("Jolly"));
            NatureTypes.Add(new NatureType("Naive"));
            NatureTypes.Add(new NatureType("Modest"));
            NatureTypes.Add(new NatureType("Mild"));
            NatureTypes.Add(new NatureType("Quiet"));
            NatureTypes.Add(new NatureType("Bashful"));
            NatureTypes.Add(new NatureType("Rash"));
            NatureTypes.Add(new NatureType("Calm"));
            NatureTypes.Add(new NatureType("Gentle"));
            NatureTypes.Add(new NatureType("Sassy"));
            NatureTypes.Add(new NatureType("Careful"));
            NatureTypes.Add(new NatureType("Quirky"));
        }
        private static void BuildMap()
        {
            BasePokemonMap.Clear();
            foreach (var pokemon in BasePokemon)
            {
                BasePokemonMap.Add(pokemon.No, pokemon);
            }
        }
        private static void ParsePokedexInfo(string[] entry) {
            // -- Pokedex description for each gen
            var pokedexInfo = new PokedexInfo {
                RedBlue = entry[37],
                Yellow = entry[38],
                Gold = entry[39],
                Silver = entry[40],
                Crystal = entry[41],
                Ruby = entry[42],
                Sapphire = entry[43]
            };

            Pokedex.Add(int.Parse(entry[0]), pokedexInfo);
        }

        private static void InsertMoveData(string[] entry, ref PokemonSpecies newPokemon) {
            string rawMoves = entry[20],
                rawMachine = entry[21],
                rawBreeding = entry[22],
                rawRby = entry[23],
                rawRbytm = entry[24],
                rawSpecial = entry[25],
                rawTutor = entry[26],
                rawAdv = entry[27],
                rawAdvTm = entry[28],
                rawAdvBreed = entry[29],
                rawAdvSpecial = entry[30],
                rawAdvTutor = entry[31],
                rawLfOnly = entry[32];
            
            newPokemon.Moves = new LearnSet
            {
                Sets = new Dictionary<MoveTypes, IReadOnlyList<Move>>
                {
                    { MoveTypes.nbGSCLevel, BuildList(rawMoves, MoveTypes.nbGSCLevel) },
                    { MoveTypes.nbGSCTM, BuildList(rawMachine, MoveTypes.nbGSCTM) },
                    { MoveTypes.nbGSCEgg, BuildList(rawBreeding, MoveTypes.nbGSCEgg) },
                    { MoveTypes.nbGSCSpecial, BuildList(rawSpecial, MoveTypes.nbGSCSpecial) },
                    { MoveTypes.nbGSCTutor, BuildList(rawTutor, MoveTypes.nbGSCTutor) },
                    { MoveTypes.nbRBYLevel, BuildList(rawRby, MoveTypes.nbRBYLevel) },
                    { MoveTypes.nbRBYTM, BuildList(rawRbytm, MoveTypes.nbRBYTM) },
                    { MoveTypes.nbAdvLevel, BuildList(rawAdv, MoveTypes.nbAdvLevel) },
                    { MoveTypes.nbAdvTM, BuildList(rawAdvTm, MoveTypes.nbAdvTM) },
                    { MoveTypes.nbAdvEgg, BuildList(rawAdvBreed, MoveTypes.nbAdvEgg) },
                    { MoveTypes.nbAdvSpecial, BuildList(rawAdvSpecial, MoveTypes.nbAdvSpecial) },
                    { MoveTypes.nbAdvTutor, BuildList(rawAdvTutor, MoveTypes.nbAdvTutor) },
                    { MoveTypes.nbAdvFL, BuildList(rawLfOnly, MoveTypes.nbAdvFL) }
                }
            };
        }

        public static List<Move> BuildList(string intString, MoveTypes type) {
            var result = new List<Move>();

            foreach (var item in ParseCsvIntString(intString)) {
                if (item == 0) continue;
                Move move = MoveDatabase.Moves[item];
                move.Source = type;
                result.Add(move);
            }

            return result;
        }

        private static void ParsePokemonDbLine(string[] entry) {
            var traitArr = new Traits[2];
            traitArr[0] = (Traits) int.Parse(entry[8]);
            traitArr[1] = (Traits) int.Parse(entry[9]);

            // -- Raw moves (Comma separated) for this poke.
            var evoArray = new int[5];
            evoArray[0] = int.Parse(entry[46]);
            evoArray[1] = int.Parse(entry[48]);
            evoArray[2] = int.Parse(entry[50]);
            evoArray[3] = int.Parse(entry[52]);
            evoArray[4] = int.Parse(entry[54]);

            var evoMArray = new int[5];
            evoMArray[0] = int.Parse(entry[47]);
            evoMArray[1] = int.Parse(entry[49]);
            evoMArray[2] = int.Parse(entry[51]);
            evoMArray[3] = int.Parse(entry[53]);
            evoMArray[4] = int.Parse(entry[55]);

            var illegalsArray = new string[4];
            illegalsArray[0] = entry[62];
            illegalsArray[1] = entry[63];
            illegalsArray[2] = entry[64];
            illegalsArray[3] = entry[65];

            var breedIllegals = new string[4];
            breedIllegals[0] = entry[66];
            breedIllegals[1] = entry[67];
            breedIllegals[2] = entry[68];
            breedIllegals[3] = entry[69];

            var rawMoveLevels = new string[4];
            rawMoveLevels[0] = entry[70];
            rawMoveLevels[1] = entry[71];
            rawMoveLevels[2] = entry[72];
            rawMoveLevels[3] = entry[73];

            var result = new PokemonSpecies {
                No = int.Parse(entry[0]),
                GSNo = int.Parse(entry[1]),
                AdvNo = int.Parse(entry[2]),
                Name = entry[3],
                Legendary = int.Parse(entry[4]) > 0,
                Uber = int.Parse(entry[5]) > 0,
                Type1 = (Elements)(int.Parse(entry[6])),
                Type2 = (Elements)(int.Parse(entry[7])),
                PAtt = traitArr,
                Color1 = int.Parse(entry[10]),
                Color2 = int.Parse(entry[11]),

                BaseSpecial = int.Parse(entry[18]),
                /* Loaded but never used in the original code. */
                //StartsWith = byte.Parse(entry[19]),
                ExistRBY = int.Parse(entry[33]) > 0,
                ExistGSC = int.Parse(entry[34]) > 0,
                ExistADV = int.Parse(entry[35]) > 0,
                PercentFemale = int.Parse(entry[36]),
                /* Loaded but unused in original source */
//                MyStage = int.Parse(entry[44]),
                
                MyMethod = int.Parse(entry[45]), // -- How this pokemon is evolved into. 3 -> Stone. 8 -> Happiness. Base forms 0. 1 = level. 2 = trade. 3 = thunder stone.
                Evo = evoArray,
                EvoM = evoMArray,
                Weight = int.Parse(entry[56]),
                Height = int.Parse(entry[57]),
                /*Offset = byte.Parse(entry[58]),
                LevelBal = byte.Parse(entry[59]),*/
                EggGroup1 = byte.Parse(entry[60]),
                EggGroup2 = byte.Parse(entry[61]),
                Illegals = illegalsArray,
                BreedIllegals = breedIllegals,
            };

            result.BaseStats = new StatSet {
                [StatTypes.nbHP] = int.Parse(entry[12]),
                [StatTypes.nbAttack] = int.Parse(entry[13]),
                [StatTypes.nbDefense] = int.Parse(entry[14]),
                [StatTypes.nbSpeed] = int.Parse(entry[15]),
                [StatTypes.nbSpAttack] = int.Parse(entry[16]),
                [StatTypes.nbSpDefense] = int.Parse(entry[17])
            };
            
            InsertMoveData(entry, ref result);


            // -- Fill in max stats for this pokemon.
            /*result.Attack = BattleSystem.GetStat(100, result.BaseAttack, 15);
            result.Defense = BattleSystem.GetStat(100, result.BaseDefense, 15);
            result.Speed = BattleSystem.GetStat(100, result.BaseSpeed, 15);
            result.SpecialAttack = BattleSystem.GetStat(100, result.BaseSAttack, 15);
            result.SpecialDefense = BattleSystem.GetStat(100, result.BaseSDefense, 15);*/
            result.MaxHP = BattleSystem.GetGenerationHp(CompatModes.nbTrueRBY, int.Parse(entry[12]), 15, 15, 100);
            result.MoveLevel = new byte[MoveDatabase.Moves.Count,3];

            for (var i = 0; i < 3; i++) { // -- can PROMISE this will break.
                for (var y = 0; y < rawMoveLevels[i].Length; y += 6) {
                    int moveTemp = int.Parse(rawMoveLevels[i].Substring(y, 3));
                    int levTemp = int.Parse(rawMoveLevels[i].Substring(y + 3, 3));
                    result.MoveLevel[moveTemp, i] = (byte)levTemp;
                }
            }

            result.ModAttr = new Traits[2];
            result.ModAttr[0] = result.PAtt[0];
            result.ModAttr[1] = result.PAtt[1];

            BasePokemon.Add(result);
        }

        public static IEnumerable<int> ParseCsvIntString(string input) {
            return string.IsNullOrEmpty(input) ? new int[0] : input.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries).Select(int.Parse);
        }

        public static bool IsLegalMove(PokemonSpecies poke, Move move, CompatModes compatMode)
        {
            List<Move> legalMoveSet = poke.GetCurrentMoveset(compatMode);
            List<int> illegalMoveIds = new List<int>();
            if (compatMode == CompatModes.nbTrueGSC)
            {
                illegalMoveIds.AddRange(poke.Illegals[0].Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries).Select(int.Parse));
                illegalMoveIds.AddRange(poke.Illegals[1].Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries).Select(int.Parse));
            }
            else
            {
                illegalMoveIds.AddRange(poke.Illegals[2].Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries).Select(int.Parse));
                illegalMoveIds.AddRange(poke.Illegals[3].Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries).Select(int.Parse));
            }
            
            if (compatMode == CompatModes.nbModAdv) // -- Illegal Moves from DB Mods
            {
                illegalMoveIds.AddRange(poke.IllegalMod.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries).Select(int.Parse));
            }
            
            return legalMoveSet.Contains(move) && !illegalMoveIds.Contains(move.ID);
        }
    }
}
