using System;
using System.Collections.Generic;
using Netbattle.Forms;
using Netbattle.Network;

namespace Netbattle.Common {
    public interface IRegPacket {
        string Command { get; }
        void Read(ByteBuffer reader);
        void Write(ByteBuffer writer);
        void Handle(ServerList listForm);
    }
    public interface IPacket {
        string Command { get; }
        void Read(ByteBuffer reader);
        void Write(ByteBuffer writer);
        void Handle(NbClient client);
    }

    public struct ServerListing {
        public int ServerNumber { get; set; }
        public string Name { get; set; }
        public string Owner { get; set; }
        public string Description { get; set; }
        public string Ip { get; set; }
        public int OnlinePlayers { get; set; }
        public int MaxPlayers { get; set; }
    }

    public delegate void EmptyEventArgs();

    public enum LogType {
        Verbose,
        Debug,
        Error,
        Warning,
        Info,
        Chat,
        Command
    }

    public struct LogItem {
        public LogType Type;
        public DateTime Time;
        public string Message;
    }

    public abstract class TaskItem {
        public TimeSpan Interval;
        public DateTime LastRun;
        public abstract void Setup();
        public abstract void Main();
        public abstract void Teardown();
    }

    public struct ServerInfo {
        public string ServerVersion;
        public int FloodTolerance;
        public int OnlinePlayers;
        public int MaxPlayers;
        public string ServerName;
    }

    public delegate void ServerInfoEventArgs(ServerInfo info);

    public delegate void PrivateMessageEventArgs(Player player, string message);
    public delegate void MessageEventArgs(string message);

    public enum Traits {
        nbNoTrait,
        nbStench,
        nbDrizzle,
        nbSpeedBoost,
        nbBattleArmor,
        nbSturdy,
        nbDamp,
        nbLimber,
        nbSandVeil,
        nbStatic,
        nbVoltAbsorb,
        nbWaterAbsorb,
        nbOblivious,
        nbCloudNine,
        nbCompoundEyes,
        nbInsomnia,
        nbColorChange,
        nbImmunity,
        nbFlashFire,
        nbShieldDust,
        nbOwnTempo,
        nbSuctionCups,
        nbIntimidate,
        nbShadowTag,
        nbRoughSkin,
        nbWonderGuard,
        nbLevitate,
        nbEffectSpore,
        nbSynchronize,
        nbClearBody,
        nbNaturalCure,
        nbLightningRod,
        nbSereneGrace,
        nbSwiftSwim,
        nbChlorophyll,
        nbIlluminate,
        nbTrace,
        nbHugePower,
        nbPoisonPoint,
        nbInnerFocus,
        nbMagmaArmor,
        nbWaterVeil,
        nbMagnetPull,
        nbSoundproof,
        nbRainDish,
        nbSandStream,
        nbPressure,
        nbThickFat,
        nbEarlyBird,
        nbFlameBody,
        nbRunAway,
        nbKeenEye,
        nbHyperCutter,
        nbPickup,
        nbTruant,
        nbHustle,
        nbCuteCharm,
        nbPlus,
        nbMinus,
        nbForecast,
        nbStickyHold,
        nbShedSkin,
        nbGuts,
        nbMarvelScale,
        nbLiquidOoze,
        nbOvergrow,
        nbBlaze,
        nbTorrent,
        nbSwarm,
        nbRockHead,
        nbDrought,
        nbArenaTrap,
        nbVitalSpirit,
        nbWhiteSmoke,
        nbPurePower,
        nbShellArmor,
        nbCacophony,
        nbAirLock
    }
    public enum Items {
        nbNoItem,
        nbBerry,
        nbBerryJuice,
        nbBitterBerry,
        nbBurntBerry,
        nbGoldBerry,
        nbIceBerry,
        nbMintBerry,
        nbMiracleBerry,
        nbMysteryBerry,
        nbParalyzecureBerry,
        nbPoisoncureBerry,
        nbBerserkGene,
        nbBlackBelt,
        nbBlackGlasses,
        nbBrightPowder,
        nbCharcoal,
        nbDragonFang,
        nbFocusBand,
        nbHardStone,
        nbKingsRock,
        nbLeftovers,
        nbLightBall,
        nbLuckyPunch,
        nbMagnet,
        nbMetalCoat,
        nbMetalPowder,
        nbMiracleSeed,
        nbMysticWater,
        nbNevermeltIce,
        nbPinkBow,
        nbPoisonBarb,
        nbPolkadotBow,
        nbQuickClaw,
        nbScopeLens,
        nbSharpBeak,
        nbSilverPowder,
        nbSoftSand,
        nbSpellTag,
        nbStick,
        nbThickClub,
        nbTwistedSpoon,
        nbCheriBerry,
        nbChestoBerry,
        nbPechaBerry,
        nbRawstBerry,
        nbAspearBerry,
        nbLeppaBerry,
        nbOranBerry,
        nbPersimBerry,
        nbLumBerry,
        nbSitrusBerry,
        nbFigyBerry,
        nbIapapaBerry,
        nbMagoBerry,
        nbWikiBerry,
        nbAguavBerry,
        nbLiechiBerry,
        nbGanlonBerry,
        nbSalacBerry,
        nbPetayaBerry,
        nbApicotBerry,
        nbLansatBerry,
        nbStarfBerry,
        nbChoiceBand,
        nbDeepSeaScale,
        nbDeepSeaTooth,
        nbLaxIncense,
        nbMachoBrace,
        nbMentalHerb,
        nbSeaIncense,
        nbShellBell,
        nbSilkScarf,
        nbSoulDew,
        nbWhiteHerb
    }
    public enum Conditions {
        nbNrm = 1,
        nbPsn,
        nbTox,
        nbSlp,
        nbBrn,
        nbPar,
        nbFrz,
        nbFnt,
    }
    public enum CompatModes {
        nbRBYTrade,
        nbGSCTrade,
        nbTrueRuSa,
        nbFullAdvance,
        nbModAdv,
        nbTrueRBY,
        nbTrueGSC,
    }

    public enum GraphicsMode {
        nbGFXGrn, // -- Green
        nbGFXRB, // -- Red/Blue
        nbGFXYlo, // -- Yellow
        nbGFXGld, // -- Gold
        nbGFXSil, // -- Silver
        nbGFXRS,// -- Ruby/Sapphire
        nbGFXLF, // -- Leafgreen/Firered
        nbGFXEme, // -- Emerald
        nbGFXCol, // -- Colosseum
        nbGFXSml, // -- small (Advance mini pics)
    }


    public struct GraphicsData {
        public long[] ByteCount { get; set; }
        public long[] ByteStart { get; set; }
        public string[] Titles { get; set; }
        public byte[] InFile { get; set; }
    }

    public enum Elements {
        nbNoType,
        nbNormal,
        nbFire,
        nbWater,
        nbElectr,
        nbGrass,
        nbIce,
        nbFight,
        nbPoison,
        nbGround,
        nbFlying,
        nbPsychc,
        nbBug,
        nbRock,
        nbGhost,
        nbDragon,
        nbDark,
        nbSteel
    }

    public enum MoveTargets {
        nbGlobal,
        nbSelectedTarget,
        nbBothEnemies,
        nbSelfAffecting,
        nbTeamAffecting,
        nbEveryoneElse,
        nbRandomEnemy,
        nbReactionTarget,
        nbMoveCalled
    }

    public class Move {
        public short ID { get; set; }
        public string Name { get; set; }
        public Elements Type { get; set; }
        public short Power { get; set; }
        public byte Accuracy { get; set; }
        public byte PP { get; set; }
        public string Text { get; set; }
        public byte SpecialPercent { get; set; }
        public byte SpecialEffect { get; set; }
        public bool WorksRight { get; set; }
        public bool BrightPowder { get; set; }
        public bool KingsRock { get; set; }
        public bool RBYMove { get; set; }
        public bool GSCMove { get; set; }
        public bool AdvMove { get; set; }
        public bool HitsTeam { get; set; }
        public bool SelfMove { get; set; }
        public string OldTM { get; set; }
        public string NewTM { get; set; }
        public string ADVTM { get; set; }
        public bool SubstituteBlocks { get; set; }
        public bool HitsAll { get; set; }
        public bool SoundMove { get; set; }
        public bool PhysMove { get; set; }
        public MoveTargets Target { get; set; }
        public bool MagicCoat { get; set; }
        public string Source { get; set; }
    }

    public class Player {
        public int Id { get; set; }
        public int GameVersion { get; set; }
        public int Picture { get; set; }
        public int GraphicsVersion { get; set; }
        public int Authority { get; set; }
        public bool ShowTeam { get; set; }
        public bool StadiumOk { get; set; }
        public List<Pokemon> Team { get; set; }
        public int Wins { get; set; }
        public int Losses { get; set; }
        public int Ties { get; set; }
        public int Disconnects { get; set; }
        public int Compatibility { get; set; }
        public int TeamPower { get; set; }
        public int BattlingWith { get; set; } // -- TODO: Make this a reference to another player?

        public int Ping { get; set; } // -- Client's ping with the server..
        public string Name { get; set; }
        public string Version { get; set; }
        public string Description { get; set; }
    }

    public struct PokedexInfo {
        public string RedBlue { get; set; }
        public string Yellow { get; set; }
        public string Gold { get; set; }
        public string Silver { get; set; }
        public string Crystal { get; set; }
        public string Ruby { get; set; }
        public string Sapphire { get; set; }
    }
    public delegate void PlayerEventArgs(Player player);
    public delegate void PlayersEventArgs(List<Player> players);
}
