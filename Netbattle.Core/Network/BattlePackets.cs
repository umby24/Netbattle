using Netbattle.Core.Common;

namespace Netbattle.Core.Network;

public struct RandomTeam : IBattlePacket {
    public string Command => "RAND";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct BattleInfo : IBattlePacket {
    public string Command => "INFO";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct StadiumPokemon : IBattlePacket {
    public string Command => "SPKM";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct BattleStartBCMD : IBattlePacket {
    public string Command => "BCMD";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct EndTurn : IBattlePacket {
    public string Command => "ENDTN";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct WatcherSync : IBattlePacket {
    public string Command => "SYNC";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct BattleCancel : IBattlePacket {
    public string Command => "CANCL";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct PlayerBackBattle : IBattlePacket {
    public string Command => "BACK";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct ChatMessageBattle : IBattlePacket {
    public string Command => "CMSG";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct WinMessageBattle : IBattlePacket {
    public string Command => "WMSG";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct NewWatcher : IBattlePacket {
    public string Command => "WTCH";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct WatcherLeft : IBattlePacket {
    public string Command => "DONW";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct KillBattleWindow : IBattlePacket {
    public string Command => "EXIT";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct PlayerWaiting : IBattlePacket {
    public string Command => "WAIT";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct ServerMessageBattle : IBattlePacket {
    public string Command => "SMSG";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct HackDetected : IBattlePacket {
    public string Command => "HACK";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct TimeoutWarning : IBattlePacket {
    public string Command => "HURRY";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct BattleTimeout : IBattlePacket {
    public string Command => "TIME";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct DumpNotification : IBattlePacket {
    public string Command => "DUMP";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct RequestUnrate : IBattlePacket {
    public string Command => "REQUN";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct UnrateRequestAccepted : IBattlePacket {
    public string Command => "UNACC";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct RequestTie : IBattlePacket {
    public string Command => "REQTI";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct TieAccepted : IBattlePacket {
    public string Command => "TIACC";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct RequestRefused : IBattlePacket {
    public string Command => "REFUS";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct BattleOver : IBattlePacket {
    public string Command => "BOVER";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}

public struct WatcherIgnoringChat : IBattlePacket {
    public string Command => "IGWAT";
    public void Read(ByteBuffer reader) {
        throw new NotImplementedException();
    }

    public void Write(ByteBuffer writer) {
        throw new NotImplementedException();
    }

    public void Handle(NbClient client) {
        throw new NotImplementedException();
    }
}
