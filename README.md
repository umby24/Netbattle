# Pokemon Netbattle
This is a C# reimplementation of the [Pokemon Netbattle](https://bulbapedia.bulbagarden.net/wiki/Pok%C3%A9mon_NetBattle) game.

# Why?
This is part of a project of mine to port the old game to a newer language.

Netbattle was a big part of my life for quite a few years. I still have the original codebase lying around, and figured I could give it some justice by bringing it to the open source world, and give myself a programming challenge to port it.

# What's complete?
- XOR Encryption
- Registry communication
- Basic chat
- Private messages
- Loading the original netbattle databases (Pokemon, Type, Move, Graphics)

# Original Bugs Discovered
- RBY would set SP. Atk and Def. Identically, and set the HP DV as 0.
- Challenge Cup teams, all pokemon always were generated with an HP DV of 0.
- Instead of Dragonite, Draonair, and dratini learning ExtremeSpeed + Twister, it was Dragonite, Moltres, and dratini. (GSC Rules)
- Obviously, the infamous 'PM to No One' crash