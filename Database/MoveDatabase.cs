using System;
using System.Collections.Generic;
using System.Linq;
using Netbattle.Common;

namespace Netbattle.Database {
    public enum MoveTypes
    {
        nbRBYLevel = 1,
        nbRBYTM = 2,
        nbGSCLevel = 3,
        nbGSCTM = 4,
        nbGSCEgg = 5,
        nbGSCTutor = 6,
        nbGSCSpecial = 7,
        nbAdvLevel = 8,
        nbAdvTM = 9,
        nbAdvEgg = 10,
        nbAdvTutor = 11,
        nbAdvSpecial = 12,
        nbAdvFL = 13
    }
    
    /// <summary>
    /// Loader and manager of the Pokemon Move Database.
    /// </summary>
    public class MoveDatabase {
        public static Dictionary<int, Move> Moves = new Dictionary<int, Move>();

        public static void Load() {
            var moveDb = new CdbFile("MoveDB.cdb");
            moveDb.Load();
            Logger.Log(LogType.Verbose, "MoveDB Read and decompressed.");

            foreach (string[] entry in moveDb.LineContent) {
                ParseFileEntry(entry);
            }

            Logger.Log(LogType.Verbose, $"MoveDB load complete. {Moves.Count} total moves.");
        }

        /// <summary>
        /// Validates if a Pokemon's moveset is legal according to game rules
        /// </summary>
        /// <param name="pokemon">The Pokemon to validate</param>
        /// <param name="ignoreDVs">Whether to ignore DV checks for Odd Egg Pokemon</param>
        /// <returns>Empty string if legal, error message if illegal</returns>
        public string LegalMove(Pokemon pokemon, bool ignoreDVs = false)
        {
            short x, y;
            short breedingMoves = 0;
            short gscBreedingMoves = 0;
            short rbyMoves = 0;
            short rbyConflict = 0;
            short gscConflict = 0;
            bool invalidMove;
            short specialMoves = 0;
            short gscSpecialMoves = 0;
            bool oddEggPoke = false;
            string temp;
            bool b = false;
            int[] mv = new int[5]; // Index 0 unused, 1-4 used
            bool[,] allMoves = new bool[Moves.Count, 14]; // Second dimension 0-13 (0 unused, 1-13 used)

            // Initialize the AllMoves array to false
            for (x = 0; x < Moves.Count; x++)
            {
                for (y = 1; y <= 13; y++)
                {
                    allMoves[x, y] = false;
                }
            }

            // Populate valid moves based on game version
            switch (pokemon.GameVersion)
            {
                case CompatModes.nbTrueRBY:
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbRBYLevel);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbRBYTM);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbGSCSpecial);
                    break;

                case CompatModes.nbTrueGSC:
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbGSCLevel);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbGSCTM);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbGSCEgg);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbGSCTutor);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbGSCSpecial);
                    break;

                case CompatModes.nbRBYTrade:
                case CompatModes.nbGSCTrade:
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbRBYLevel);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbRBYTM);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbGSCLevel);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbGSCTM);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbGSCEgg);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbGSCTutor);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbGSCSpecial);
                    break;

                case CompatModes.nbTrueRuSa:
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbAdvLevel);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbAdvTM);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbAdvEgg);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbAdvSpecial);
                    break;

                case CompatModes.nbFullAdvance:
                case CompatModes.nbModAdv:
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbAdvLevel);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbAdvTM);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbAdvEgg);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbAdvSpecial);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbAdvTutor);
                    ValidMoveArray(pokemon, allMoves, MoveTypes.nbAdvFL);
                    break;

                default:
                    return "Invalid Input";
            }

            // Clear the 0 index
            for (x = 1; x <= 13; x++)
            {
                allMoves[0, x] = false;
            }

            // Copy moves to local array
            for (x = 1; x <= 4; x++)
            {
                mv[x] = pokemon.Move[x];
            }

            invalidMove = false;
            breedingMoves = 0;
            gscBreedingMoves = 0;
            rbyMoves = 0;

            // Check 1: Invalid moves - can the Pokemon learn this move at all?
            for (x = 1; x <= 4; x++)
            {
                if (mv[x] != 0)
                {
                    invalidMove = true;
                    for (y = 1; y <= 13; y++)
                    {
                        if (allMoves[mv[x], y])
                        {
                            invalidMove = false;
                            break;
                        }
                    }
                    if (invalidMove)
                    {
                        return $"{pokemon.Name} can't learn {Moves[mv[x]].Name}.  There may have been a recent change to the database.";
                    }
                }
            }

            // Check 2: Database-defined illegal combinations
            if (CompatVersion(pokemon.GameVersion) == BattleModes.nbGSCBattle)
            {
                temp = pokemon.Illegals[0];
                if (pokemon.GameVersion == CompatModes.nbTrueGSC)
                {
                    temp += pokemon.Illegals[1];
                }
            }
            else
            {
                temp = pokemon.Illegals[2];
                if (pokemon.GameVersion == CompatModes.nbTrueGSC)
                {
                    temp += pokemon.Illegals[3];
                }
            }

            if (!DBIllegalCheck(mv, temp))
            {
                return $"{pokemon.Name}'s moveset contains a combination of moves that is not legally obtainable.";
            }

            if (pokemon.GameVersion == CompatModes.nbModAdv)
            {
                if (!DBIllegalCheck(mv, pokemon.IllegalMod))
                {
                    return $"The currently loaded Database Mod prohibits this moveset on {pokemon.Name}";
                }
            }

            // Check 3: Invalid breeding combinations
            if (!BreedCheck(pokemon.No, mv, pokemon.GameVersion))
            {
                return $"{pokemon.Name}'s moveset contains a combination of Egg Moves that is not legally obtainable.";
            }

            // Check 4: RBY Moves <-/-> GSC Breed, GSC Special
            rbyMoves = 0;
            for (x = 1; x <= 4; x++)
            {
                if (LegalMoveCheck(allMoves, (short)mv[x], MoveTypes.nbRBYLevel, MoveTypes.nbRBYTM))
                {
                    rbyMoves = (short)mv[x];
                    break;
                }
            }

            rbyConflict = 0;
            for (x = 1; x <= 4; x++)
            {
                if (LegalMoveCheck(allMoves, (short)mv[x], MoveTypes.nbGSCEgg, MoveTypes.nbGSCSpecial))
                {
                    if (!Moves[mv[x]].RBYMove)
                    {
                        rbyConflict = (short)mv[x];
                        b = false;
                        break;
                    }
                }
            }

            if (rbyConflict != 0 && rbyMoves != 0)
            {
                temp = $"{pokemon.Name} cannot learn both {Moves[rbyMoves].Name} and {Moves[rbyConflict].Name}.\n";
                temp += "(Cannot combine RBY Moves and non-RBY Breeding or Special Moves.)";
                return temp;
            }

            // Check 5: Only 1 Special Move (with exceptions)
            specialMoves = 0;
            y = 0;
            for (x = 1; x <= 4; x++)
            {
                if (LegalMoveCheck(allMoves, (short)mv[x], MoveTypes.nbAdvSpecial, MoveTypes.nbGSCSpecial))
                {
                    if (y == 0)
                    {
                        y = (short)mv[x];
                    }
                    else
                    {
                        specialMoves = (short)mv[x];
                        break;
                    }
                }
            }

            if (specialMoves != 0)
            {
                if (CompatVersion(pokemon.GameVersion) == BattleModes.nbAdvBattle)
                {
                    switch (pokemon.No)
                    {
                        case 96: // DROWZEE
                        case 97: // HYPNO - hatched with BELLY DRUM (14) and WISH (353)
                            if ((specialMoves == 14 && y == 353) || (specialMoves == 353 && y == 14))
                            {
                                specialMoves = 0;
                            }
                            break;

                        case 108: // LICKITUNG - hatched with HEAL BELL (89) and WISH (353)
                            if ((specialMoves == 89 && y == 353) || (specialMoves == 353 && y == 89))
                            {
                                specialMoves = 0;
                            }
                            break;

                        case 113: // CHANSEY
                        case 242: // BLISSEY - hatched with SWEET SCENT (220) and WISH (353)
                            if ((specialMoves == 220 && y == 353) || (specialMoves == 353 && y == 220))
                            {
                                specialMoves = 0;
                            }
                            break;

                        case 102: // EXEGGCUTE
                        case 103: // EXEGGUTOR - hatched with SWEET SCENT (220) and WISH (353)
                            if ((specialMoves == 220 && y == 353) || (specialMoves == 353 && y == 220))
                            {
                                specialMoves = 0;
                            }
                            break;

                        case 115: // KANGASKHAN - hatched with YAWN (354) and WISH (353)
                            if ((specialMoves == 354 && y == 353) || (specialMoves == 353 && y == 354))
                            {
                                specialMoves = 0;
                            }
                            break;

                        case 83: // FARFETCH'D - hatched with YAWN (354) and WISH (353)
                            if ((specialMoves == 354 && y == 353) || (specialMoves == 353 && y == 354))
                            {
                                specialMoves = 0;
                            }
                            break;
                    }
                }
            }

            if (specialMoves != 0)
            {
                temp = $"{pokemon.Name} cannot learn both {Moves[y].Name} and {Moves[specialMoves].Name}.\n";
                temp += "(Cannot combine two or more Special Moves.)";
                return temp;
            }

            // Check 6: Egg Moves <-/-> Special Moves
            breedingMoves = 0;

            // EXCEPTION: Gligar can have Earthquake and {Wing Attack and/or Counter}
            b = false;
            if (pokemon.No == 207) // Gligar
            {
                for (x = 1; x <= 4; x++)
                {
                    if (mv[x] == 55) // Earthquake
                    {
                        b = true;
                        break;
                    }
                }
            }

            specialMoves = 0;
            for (x = 1; x <= 4; x++)
            {
                if (LegalMoveCheck(allMoves, (short)mv[x], MoveTypes.nbGSCSpecial, MoveTypes.nbAdvSpecial))
                {
                    specialMoves = (short)mv[x];
                    break;
                }
            }

            for (x = 1; x <= 4; x++)
            {
                if (!Moves[specialMoves].RBYMove)
                {
                    if (LegalMoveCheck(allMoves, (short)mv[x], MoveTypes.nbGSCEgg, MoveTypes.nbAdvEgg, ignoreRBY: true))
                    {
                        if ((mv[x] != 248 && mv[x] != 34) || !b) // Not Counter (68) or Wing Attack (17) with Gligar exception
                        {
                            breedingMoves = (short)mv[x];
                            break;
                        }
                    }
                }
                else
                {
                    if (LegalMoveCheck(allMoves, (short)mv[x], MoveTypes.nbGSCEgg, MoveTypes.nbAdvEgg))
                    {
                        if ((mv[x] != 248 && mv[x] != 34) || !b)
                        {
                            breedingMoves = (short)mv[x];
                            break;
                        }
                    }
                }
            }

            if (breedingMoves != 0 && specialMoves != 0)
            {
                temp = $"{pokemon.Name} cannot learn both {Moves[breedingMoves].Name} and {Moves[specialMoves].Name}.\n";
                temp += "(Cannot combine Egg Moves with Special Moves.)";
                return temp;
            }

            if (CompatVersion(pokemon.GameVersion) != BattleModes.nbAdvBattle)
            {
                // Check 7: GSC Egg Moves <-/-> Special Moves or RBY Moves
                gscBreedingMoves = 0;
                for (x = 1; x <= 4; x++)
                {
                    if (LegalMoveCheck(allMoves, (short)mv[x], MoveTypes.nbGSCEgg))
                    {
                        if (!Moves[mv[x]].RBYMove)
                        {
                            gscBreedingMoves = (short)mv[x];
                            break;
                        }
                    }
                }

                gscConflict = 0;
                for (x = 1; x <= 4; x++)
                {
                    if (LegalMoveCheck(allMoves, (short)mv[x], MoveTypes.nbRBYLevel, MoveTypes.nbRBYTM))
                    {
                        gscConflict = (short)mv[x];
                        break;
                    }
                }

                if (gscBreedingMoves != 0 && gscConflict != 0)
                {
                    temp = $"{pokemon.Name} cannot learn both {Moves[gscBreedingMoves].Name} and {Moves[gscConflict].Name}.\n";
                    temp += "(Cannot combine non-RBY Egg Moves with RBY Moves or Special Moves.)";
                    return temp;
                }

                // Check 8: Odd Eggs and Dizzy Punch
                switch (pokemon.No)
                {
                    case 173: // Cleffa
                    case 35:  // Clefairy
                    case 36:  // Clefable
                    case 174: // Igglybuff
                    case 39:  // Jigglypuff
                    case 40:  // Wigglytuff
                    case 236: // Tyrogue
                    case 106: // Hitmonlee
                    case 107: // Hitmonchan
                    case 175: // Togepi
                    case 176: // Togetic
                    case 237: // Hitmontop
                    case 238: // Smoochum
                    case 124: // Jynx
                    case 240: // Magby
                    case 126: // Magmar
                    case 25:  // Pikachu
                    case 26:  // Raichu
                    case 125: // Electabuzz
                    case 172: // Pichu
                    case 239: // Elekid
                        for (x = 1; x <= 4; x++)
                        {
                            if (mv[x] == 45) // Dizzy Punch
                            {
                                oddEggPoke = true;
                                break;
                            }
                        }
                        break;
                }

                // Exception: These Pokemon can't have Dizzy Punch from Odd Egg if traded from GSC
                switch (pokemon.No)
                {
                    case 25:  // Pikachu
                    case 26:  // Raichu
                    case 125: // Electabuzz
                    case 172: // Pichu
                    case 239: // Elekid
                        if (pokemon.GameVersion == CompatModes.nbGSCTrade)
                        {
                            oddEggPoke = false;
                        }
                        break;
                }

                if (oddEggPoke && breedingMoves != 0)
                {
                    temp = $"{pokemon.Name} cannot learn both Dizzy Punch and {Moves[breedingMoves].Name}.\n";
                    temp += "(Cannot combine Dizzy Punch and Egg Moves on Odd Egg Pokémon.)";
                    return temp;
                }

                if (!ignoreDVs)
                {
                    if (oddEggPoke && 
                        !((pokemon.DV_Atk == 2 && pokemon.DV_Def == 10 && pokemon.DV_SAtk == 10 && pokemon.DV_Spd == 10) ||
                          (pokemon.DV_Atk == 0 && pokemon.DV_Def == 0 && pokemon.DV_SAtk == 0 && pokemon.DV_Spd == 0)))
                    {
                        return $"DVs must be either 2/10/10/10 or 0/0/0/0 in order for Dizzy Punch to be on {pokemon.Name}";
                    }
                }
            }

            // Check 9: GSC Pre-evolution moves <-/-> RBY Moves
            rbyConflict = 0;
            switch (pokemon.No)
            {
                case 35: // Clefairy
                case 36: // Clefable
                case 39: // Jigglypuff
                case 40: // Wigglytuff
                    for (x = 1; x <= 4; x++)
                    {
                        if (mv[x] == 25 || mv[x] == 100 || mv[x] == 219) // Mega Punch, Teleport, Safeguard
                        {
                            rbyConflict = (short)mv[x];
                            break;
                        }
                    }
                    break;

                case 25: // Pikachu
                case 26: // Raichu
                case 124: // Jynx
                    for (x = 1; x <= 4; x++)
                    {
                        if (mv[x] == 25 || mv[x] == 219) // Mega Punch, Safeguard
                        {
                            rbyConflict = (short)mv[x];
                            break;
                        }
                    }
                    break;

                case 130: // Gyarados
                    for (x = 1; x <= 4; x++)
                    {
                        if (mv[x] == 68) // Counter
                        {
                            rbyConflict = (short)mv[x];
                            break;
                        }
                    }
                    break;
            }

            if (rbyMoves != 0 && rbyConflict != 0)
            {
                temp = $"{pokemon.Name} cannot learn both {Moves[rbyMoves].Name} and {Moves[rbyConflict].Name}.\n";
                temp += "(Cannot combine GSC Pre-Evolution Moves with RBY Moves.)";
                return temp;
            }

            // Check 10: Marill/Azumarill - cannot mix Marill egg moves with Azurill egg moves
            if (pokemon.No == 183 || pokemon.No == 184) // Marill or Azumarill
            {
                y = 0;
                for (x = 1; x <= 4; x++)
                {
                    switch (mv[x])
                    {
                        case 343: // Tickle
                        case 187: // Belly Drum
                        case 183: // Mach Punch
                        case 319: // Sing (from Azurill)
                        case 58:  // Ice Beam (from Azurill) 
                            y = (short)mv[x];
                            break;
                    }
                }

                x = 0;
                for (short i = 1; i <= 4; i++)
                {
                    if (LegalMoveCheck(allMoves, (short)mv[i], MoveTypes.nbAdvEgg) && y > 0)
                    {
                        switch (mv[i])
                        {
                            case 343: // Tickle
                            case 187: // Belly Drum
                            case 183: // Mach Punch
                            case 319: // Sing
                            case 58:  // Ice Beam
                                // These are Azurill moves, skip
                                break;
                            default:
                                x = (short)mv[i];
                                break;
                        }
                        if (x != 0) break;
                    }
                }

                if (x != 0 && y != 0)
                {
                    temp = $"{pokemon.Name} cannot learn both {Moves[x].Name} and {Moves[y].Name}.\n";
                    temp += "(Cannot combine Marill's Egg Moves with Azurill's Egg Moves.)";
                    return temp;
                }
            }

            // If we made it here, the moveset is legal
            return "";
        }

        /// <summary>
        /// Checks if a move is exclusively learned through specified methods
        /// </summary>
        private bool LegalMoveCheck(bool[,] allMoves, short moveNum, MoveTypes c1, MoveTypes c2 = 0, 
            MoveTypes c3 = 0, MoveTypes c4 = 0, MoveTypes c5 = 0, MoveTypes c6 = 0, 
            MoveTypes c7 = 0, MoveTypes c8 = 0, MoveTypes c9 = 0, bool ignoreRBY = false)
        {
            bool build = false;

            try
            {
                // Check if the move exists in any of the specified categories
                if (allMoves[moveNum, (int)c1]) build = true;
                if (c2 > 0 && allMoves[moveNum, (int)c2]) build = true;
                if (c3 > 0 && allMoves[moveNum, (int)c3]) build = true;
                if (c4 > 0 && allMoves[moveNum, (int)c4]) build = true;
                if (c5 > 0 && allMoves[moveNum, (int)c5]) build = true;
                if (c6 > 0 && allMoves[moveNum, (int)c6]) build = true;
                if (c7 > 0 && allMoves[moveNum, (int)c7]) build = true;
                if (c8 > 0 && allMoves[moveNum, (int)c8]) build = true;
                if (c9 > 0 && allMoves[moveNum, (int)c9]) build = true;

                // Check that the move is NOT in any other categories (exclusivity check)
                for (int x = 1; x <= 13; x++)
                {
                    if (x != (int)c1 && x != (int)c2 && x != (int)c3 && x != (int)c4 && 
                        x != (int)c5 && x != (int)c6 && x != (int)c7 && x != (int)c8 && x != (int)c9)
                    {
                        if (allMoves[moveNum, x])
                        {
                            // If ignoreRBY is true, don't fail on RBY moves
                            if (!(ignoreRBY && (x == (int)MoveTypes.nbRBYLevel || x == (int)MoveTypes.nbRBYTM)))
                            {
                                build = false;
                                break;
                            }
                        }
                    }
                }
            }
            catch
            {
                // Error handling - return false on any exception
                build = false;
            }

            return build;
        }

        /// <summary>
        /// Populates the valid moves array for a specific move type category
        /// </summary>
        private void ValidMoveArray(Pokemon pkmn, bool[,] fillArray, MoveTypes moveType)
        {
            Pokemon basePkmn = PokemonDatabase.BasePokemon[pkmn.No];

            switch (moveType)
            {
                case MoveTypes.nbRBYLevel:
                    if (basePkmn.RBYMoves != null)
                    {
                        foreach (var move in basePkmn.RBYMoves)
                        {
                            fillArray[move.ID, (int)moveType] = true;
                        }
                    }
                    break;

                case MoveTypes.nbRBYTM:
                    if (basePkmn.RBYTM != null)
                    {
                        foreach (var move in basePkmn.RBYTM)
                        {
                            fillArray[move.ID, (int)moveType] = true;
                        }
                    }
                    break;

                case MoveTypes.nbGSCLevel:
                    if (basePkmn.BaseMoves != null)
                    {
                        foreach (var move in basePkmn.BaseMoves)
                        {
                            fillArray[move.ID, (int)moveType] = true;
                        }
                    }
                    break;

                case MoveTypes.nbGSCTM:
                    if (basePkmn.MachineMoves != null)
                    {
                        foreach (var move in basePkmn.MachineMoves)
                        {
                            fillArray[move.ID, (int)moveType] = true;
                        }
                    }
                    break;

                case MoveTypes.nbGSCEgg:
                    if (basePkmn.BreedingMoves != null)
                    {
                        foreach (var move in basePkmn.BreedingMoves)
                        {
                            fillArray[move.ID, (int)moveType] = true;
                        }
                    }
                    break;

                case MoveTypes.nbGSCTutor:
                    if (basePkmn.MoveTutor != null)
                    {
                        foreach (var move in basePkmn.MoveTutor)
                        {
                            fillArray[move.ID, (int)moveType] = true;
                        }
                    }
                    break;

                case MoveTypes.nbGSCSpecial:
                    // Only fill special moves if NOT in True RBY or True GSC mode
                    if (pkmn.GameVersion != CompatModes.nbTrueGSC && pkmn.GameVersion != CompatModes.nbTrueRBY)
                    {
                        if (basePkmn.SpecialMoves != null)
                        {
                            foreach (var move in basePkmn.SpecialMoves)
                            {
                                fillArray[move.ID, (int)moveType] = true;
                            }
                        }
                    }
                    break;

                case MoveTypes.nbAdvLevel:
                    if (basePkmn.AdvMoves != null)
                    {
                        int limit = (pkmn.GameVersion == CompatModes.nbModAdv) ? 
                            basePkmn.AdvMoves.Count : (int)basePkmn.TotalAdvMoves;
                        for (int i = 0; i < limit && i < basePkmn.AdvMoves.Count; i++)
                        {
                            fillArray[basePkmn.AdvMoves[i].ID, (int)moveType] = true;
                        }
                    }
                    break;

                case MoveTypes.nbAdvTM:
                    if (basePkmn.ADVTM != null)
                    {
                        foreach (var move in basePkmn.ADVTM)
                        {
                            fillArray[move.ID, (int)moveType] = true;
                        }
                    }
                    break;

                case MoveTypes.nbAdvEgg:
                    if (basePkmn.AdvBreeding != null)
                    {
                        foreach (var move in basePkmn.AdvBreeding)
                        {
                            fillArray[move.ID, (int)moveType] = true;
                        }
                    }
                    break;

                case MoveTypes.nbAdvTutor:
                    if (basePkmn.AdvTutor != null)
                    {
                        foreach (var move in basePkmn.AdvTutor)
                        {
                            fillArray[move.ID, (int)moveType] = true;
                        }
                    }
                    break;

                case MoveTypes.nbAdvSpecial:
                    if (basePkmn.AdvSpecial != null)
                    {
                        foreach (var move in basePkmn.AdvSpecial)
                        {
                            fillArray[move.ID, (int)moveType] = true;
                        }
                    }
                    break;

                case MoveTypes.nbAdvFL:
                    if (basePkmn.LFOnly != null)
                    {
                        foreach (var move in basePkmn.LFOnly)
                        {
                            fillArray[move.ID, (int)moveType] = true;
                        }
                    }
                    break;
            }

            // Filter moves based on game version
            switch (pkmn.GameVersion)
            {
                case CompatModes.nbRBYTrade:
                case CompatModes.nbTrueRBY:
                    // RBY Moves Only
                    for (int x = 1; x < Moves.Count; x++)
                    {
                        if (!Moves[x].RBYMove)
                        {
                            fillArray[x, (int)moveType] = false;
                        }
                    }
                    // Add special RBY moves for specific Pokemon
                    if (moveType == MoveTypes.nbGSCSpecial)
                    {
                        switch (basePkmn.No)
                        {
                            case 54: // Psyduck
                            case 55: // Golduck
                                fillArray[6, (int)moveType] = true; // Pay Day
                                break;
                        }
                    }
                    break;

                case CompatModes.nbGSCTrade:
                case CompatModes.nbTrueGSC:
                    // RBY/GSC Moves Only
                    for (int x = 1; x < Moves.Count; x++)
                    {
                        if (!Moves[x].GSCMove)
                        {
                            fillArray[x, (int)moveType] = false;
                        }
                    }
                    // Add special GSC moves for specific Pokemon
                    if (moveType == MoveTypes.nbGSCSpecial)
                    {
                        switch (basePkmn.No)
                        {
                            case 83: // Farfetch'd
                                fillArray[12, (int)moveType] = true; // Pay Day (GSC event)
                                break;
                            case 147: // Dratini
                            case 148: // Dragonair
                            case 149: // Dragonite
                                fillArray[61, (int)moveType] = true; // ExtremeSpeed
                                break;
                            case 207: // Gligar
                                fillArray[55, (int)moveType] = true; // Earthquake
                                break;
                            case 25:  // Pikachu
                            case 26:  // Raichu
                            case 35:  // Clefairy
                            case 36:  // Clefable
                            case 39:  // Jigglypuff
                            case 40:  // Wigglytuff
                            case 106: // Hitmonlee
                            case 107: // Hitmonchan
                            case 125: // Electabuzz
                            case 126: // Magmar
                            case 135: // Jolteon
                            case 172: // Pichu
                            case 173: // Cleffa
                            case 174: // Igglybuff
                            case 236: // Tyrogue
                            case 237: // Hitmontop
                            case 238: // Smoochum
                            case 239: // Elekid
                                fillArray[45, (int)moveType] = true; // Dizzy Punch (Odd Egg)
                                break;
                        }
                    }
                    break;

                case CompatModes.nbTrueRuSa:
                case CompatModes.nbFullAdvance:
                case CompatModes.nbModAdv:
                    // No filtering needed for Advance games
                    break;
            }
        }

        /// <summary>
        /// Determines the battle mode based on the compatibility version
        /// </summary>
        private BattleModes CompatVersion(CompatModes gameVersion)
        {
            switch (gameVersion)
            {
                case CompatModes.nbTrueRBY:
                case CompatModes.nbTrueGSC:
                case CompatModes.nbRBYTrade:
                case CompatModes.nbGSCTrade:
                    return BattleModes.nbGSCBattle;

                case CompatModes.nbTrueRuSa:
                case CompatModes.nbFullAdvance:
                case CompatModes.nbModAdv:
                    return BattleModes.nbAdvBattle;

                default:
                    return BattleModes.nbGSCBattle;
            }
        }

        /// <summary>
        /// Checks if a moveset violates database-defined illegal combinations
        /// </summary>
        private bool DBIllegalCheck(int[] moves, string illegals)
        {
            if (string.IsNullOrEmpty(illegals))
            {
                return true;
            }

            var legality = true;
            var sets = illegals.Split('|');
            foreach (var set in sets) {
                var tempSet = set.Split('+');
                legality = HasMoves(moves, tempSet.Select(short.Parse).ToList());
                if (legality)
                    break;
            }
            
            return !legality; // If any illegal set is found, return false
        }
        
        private bool HasMoves(int[] moves, List<short> checkMoves) {
            int instances = 0;
            foreach (var move in checkMoves)
            {
                foreach (var m in moves)
                {
                    if (m == move)
                    {
                        instances++;
                    }
                }
            }
            
            return instances == checkMoves.Count;
        }
        
        /// <summary>
        /// Checks if egg move combinations are legal for breeding
        /// </summary>
        private bool BreedCheck(int pokemonNo, int[] moves, CompatModes gameVersion)
        {
            // Implementation depends on breeding chain data structure
            var basePkmn = PokemonDatabase.BasePokemon[pokemonNo];
            string temp;
            if (CompatVersion(gameVersion) == BattleModes.nbGSCBattle) {
                temp = basePkmn.BreedIllegals[0];
                if (gameVersion == CompatModes.nbTrueGSC) temp += basePkmn.BreedIllegals[1];
            }
            else {
                temp = basePkmn.BreedIllegals[2];
                if (gameVersion == CompatModes.nbTrueGSC) temp += basePkmn.BreedIllegals[3];
            }
            
            return DBIllegalCheck(moves, temp);
        }
        
        private static void ParseFileEntry(IReadOnlyList<string> entry) {
            var result = new Move {
                ID = short.Parse(entry[0]),
                Name = entry[1],
                Type = (Elements)int.Parse(entry[2]),
                Power = short.Parse(entry[3]),
                Accuracy = byte.Parse(entry[4]),
                PP = byte.Parse(entry[5]),
                SpecialPercent = byte.Parse(entry[6]),
                SpecialEffect = byte.Parse(entry[7]),
                Target = (MoveTargets)int.Parse(entry[8]),
                Text = entry[9],
                WorksRight = int.Parse(entry[10]) > 0,
                BrightPowder = int.Parse(entry[11]) > 0,
                KingsRock = int.Parse(entry[12]) > 0,
                RBYMove = int.Parse(entry[13]) > 0,
                GSCMove = int.Parse(entry[14]) > 0,
                AdvMove = int.Parse(entry[15]) > 0,
                HitsTeam = int.Parse(entry[16]) > 0,
                SelfMove = int.Parse(entry[17]) > 0,
                OldTM = entry[18],
                NewTM = entry[19],
                ADVTM = entry[20],
                SubstituteBlocks = int.Parse(entry[21]) > 0,
                HitsAll = int.Parse(entry[22]) > 0,
                SoundMove = int.Parse(entry[23]) > 0,
                PhysMove = int.Parse(entry[24]) > 0,
                MagicCoat = int.Parse(entry[25]) > 0
            };

            Moves.Add(result.ID, result);
        }
    }
}
