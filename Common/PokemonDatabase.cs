using System;
using System.Collections.Generic;
using System.Linq;
using Netbattle.Database;

namespace Netbattle.Common {
    public class PokemonDatabase {
        public static List<Pokemon> BasePokemon = new List<Pokemon>();
        public static Dictionary<int, PokedexInfo> Pokedex = new Dictionary<int, PokedexInfo>();

        public static void Load() {
            var pokeDb = new CdbFile("PokeDB.cdb");
            pokeDb.Load();
            Logger.Log(LogType.Verbose, "PokemonDB Read and decompressed.");

            foreach (string[] entry in pokeDb.LineContent) {
                ParsePokedexInfo(entry);
                ParsePokemonDbLine(entry);
            }

            Logger.Log(LogType.Info, $"Pokemon Database loaded successfully. Showing {BasePokemon.Count} Pokemon.");
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

        private static void InsertMoveData(string[] entry, ref Pokemon newPokemon) {
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

            newPokemon.BaseMoves = BuildList(rawMoves, "Level");
            newPokemon.MachineMoves = BuildList(rawMachine, "Meh");
            newPokemon.BreedingMoves = BuildList(rawBreeding, "Egg Move");
            newPokemon.RBYMoves = BuildList(rawRby, "Level");
            newPokemon.RBYTM = BuildList(rawRbytm, "RBYTM");
            newPokemon.SpecialMoves = BuildList(rawSpecial, "Special");
            newPokemon.AdvMoves = BuildList(rawAdv, "Level");
            newPokemon.ADVTM = BuildList(rawAdvTm, "TM");
            newPokemon.AdvBreeding = BuildList(rawAdvBreed, "Egg Move");
            newPokemon.AdvSpecial = BuildList(rawAdvSpecial, "Box/NYPC");
            newPokemon.AdvTutor = BuildList(rawAdvTutor, "Move Tutor");
            newPokemon.LFOnly = BuildList(rawLfOnly, "Fire/Leaf");
            newPokemon.MoveTutor = new List<Move>();

            // -- Not really sure what this does.
            if (rawTutor.Length > 0) {
                int value = int.Parse(rawTutor);
                if (value - 4 >= 0) {
                    newPokemon.MoveTutor.Add(MoveDatabase.Moves[70]);
                    value -= 4;
                }

                if (value - 2 >= 0) {
                    newPokemon.MoveTutor.Add(MoveDatabase.Moves[98]);
                    value -= 2;
                }

                if (value - 1 >= 0) {
                    newPokemon.MoveTutor.Add(MoveDatabase.Moves[232]);
                }
            }

            newPokemon.TotalAdvMoves = newPokemon.AdvMoves.Count;
        }

        public static List<Move> BuildList(string intString, string source) {
            var result = new List<Move>();

            foreach (var item in ParseCsvIntString(intString)) {
                Move move = MoveDatabase.Moves[item];
                move.Source = source;

                if (source == "TM")
                    move.Source = move.ADVTM;

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

            var result = new Pokemon {
                No = int.Parse(entry[0]),
                GSNo = int.Parse(entry[1]),
                AdvNo = int.Parse(entry[2]),
                Name = entry[3],
                Nickname = entry[3],
                Legendary = int.Parse(entry[4]) > 0,
                Uber = int.Parse(entry[5]) > 0,
                Type1 = (Elements)(int.Parse(entry[6])),
                Type2 = (Elements)(int.Parse(entry[7])),
                PAtt = traitArr,
                Color1 = int.Parse(entry[10]),
                Color2 = int.Parse(entry[11]),
                BaseHP = int.Parse(entry[12]),
                BaseAttack = int.Parse(entry[13]),
                BaseDefense = int.Parse(entry[14]),
                BaseSpeed = int.Parse(entry[15]),
                BaseSAttack = int.Parse(entry[16]),
                BaseSDefense = int.Parse(entry[17]),
                BaseSpecial = int.Parse(entry[18]),
                StartsWith = byte.Parse(entry[19]),
                ExistRBY = int.Parse(entry[33]) > 0,
                ExistGSC = int.Parse(entry[34]) > 0,
                ExistAdv = int.Parse(entry[35]) > 0,
                PercentFemale = int.Parse(entry[36]),
                MyStage = int.Parse(entry[44]),
                MyMethod = int.Parse(entry[45]),
                Evo = evoArray,
                EvoM = evoMArray,
                Weight = int.Parse(entry[56]),
                Height = int.Parse(entry[57]),
                Offset = byte.Parse(entry[58]),
                LevelBal = byte.Parse(entry[59]),
                EggGroup1 = byte.Parse(entry[60]),
                EggGroup2 = byte.Parse(entry[61]),
                Illegals = illegalsArray,
                BreedIllegals = breedIllegals,
            };
            
            InsertMoveData(entry, ref result);


            // -- Fill in max stats for this pokemon.
            result.Attack = BattleSystem.GetStat(100, result.BaseAttack, 15);
            result.Defense = BattleSystem.GetStat(100, result.BaseDefense, 15);
            result.Speed = BattleSystem.GetStat(100, result.BaseSpeed, 15);
            result.SpecialAttack = BattleSystem.GetStat(100, result.BaseSAttack, 15);
            result.SpecialDefense = BattleSystem.GetStat(100, result.BaseSDefense, 15);
            result.MaxHP = BattleSystem.GetHp(100, result.BaseHP, 15);
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
        
    }
}
