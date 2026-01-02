Pokémon NetBattle
Readme.txt for 0.9.4
---------------------

How to Play
---------------------

Until I have time to write up a full tutorial here, please see the web site at:
http://www.netbattle.net

Changes since 0.9.3
---------------------
Fixed bugs 
	RBY: Fly and Dig now deplete PP properly. 
	RBY: Using Bide no longer crashes the game. 
	RBY: Using Disable no longer crashes the game.
	RBY: Albiet still completely useless, Whirlwind now has an accuracy of 84.4%. 
	GSC: Encoring Mirror Move or Encore no longer hangs the game. 
	GSC: Lovely Kiss has 10 base PP. 
	GSC: Sandstorm fails if it is already in effect. 
	GSCRSE: Solarbeam now does half damage during Rain Dance, Sandstorm, and Hail. 
	GSCRSE: Moonlight, Morning Sun, and Synthesis have their effectiveness reduced during Sandstorm and Hail. 
	GSCRSE: Future Sight and Doom Desire no longer count towards Bide damage. 
	GSCRSE: OKHOs now do damage equal to the target's Maximum HP as opposed to Current HP. 
	RSE: Swagger and Flatter now check the target for Own Tempo instead of the user. 
	RSE: Cute Charm can no longer activate if a Contact move causes the trait holder to faint. 
	RSE: In Doubles, Perish Song no longer gives Fainted Pokemon a Perish Count. 
	RSE: Pokemon with the trait Guts are now immune to the Attack-halving effect of Burns. 
	RSE: Using Dig/Fly with your last Pokemon in Doubles doesn't show your partner's moveset. 
	RSE: Present works correctly now. 
	All: Database changes have been implemented.  Lots of em. 
	All: Move failure due to type immunity no longer produces "But it failed!" 
	Server: User Databases can now handle more than 35676 users. 
	Server: Tempbans no longer become Permbans if the server crashes 
	General: Blank lines no longer increase floodcount when using Multiline Paste. 
	General: Blank Win/Loss Messages no longer appear when the battle ends due to Timeout 
    
New Features 
	Server User Count now only counts different users, not clones. 
	Tutors and Deoxys from Pokemon Emerald have been implemented 
	Alias Search for mods. 
	Pokemon Nicknames can now be up to 15 characters in length. 
	Option to auto-load the most recently used team on startup. 
	Server Names can be regestered to prevent name theft. 
	Main chat window finally has a text limit 
	IPs can now be banned by range


Recent History
---------------------
0.9.3 (Changes from 0.9.2 -> 0.9.3)
	Advance mode added
	Help system added (Not finished yet)
	Colosseum & Fire/Leaf sprites added
	Fixes to GSC and RBY mode (should be 100% accurate now)
	Modifications to the scripting system
	Changed challenge windows to accomodate the new rules
	Team Builder changes (The Version menu is important!)
	New unified DataDex
	Server window and main window remember their last size
	Fancy Text is locked to On to fix the scrolling bug
	If Pokemon change when you change Team Builder versions, a list of changes is generated
	Challenge Window remembers your last selected mode for each generation, plus the last terrain selected. 
	Versions updated - Deoxys & Leaf/Fire removed from 202 mode (Now Ru/Sa Only), NYPC moves removed from True GSC, Stadium moves,
	Odd Egg moves, Box moves, etc. available in "True" modes.
	RBY Challenge Cup = True RBY, Leaf/Fire moves added to Advance Challenge Cup, Shedinja will NOT be picked in Challenge Cup.
	The installer can now download & install the Colosseum sprites (requires an active Internet connection while installing).
	Ability to show version icons and color incompatible players on the server (see the Options menu in the server window) 
0.9.2H (Test Release)
	Flash Fire now powers up only Fire moves 
	Metronome can call all moves instead of just Fly (sorry >_<) 
 	Transform/Speed issue fixed 
 	Egg Group searching now works 
	Snore acts correctly when called by a MoveCaller 
	Stadium Team Builder is not accessable in battle. 
	Stadium Mode Choose window now shows the info for the first Poke instead of the template 
	Stadium Mode shows three status PokeBalls instead of six. 
	Azurill shows up in Marill's and Azumarill's Evo Trees 
	Banning the same name or a blank string is now impossible 
	The hover texts for 0 accuracy moves are now fixed 
	Switching versions with a shiny poke now updates the image correctly 
	The Timestamp display in the Options window is fixed 
	New Wrap effects don't override old ones 
	Hi Jump Kick's accuracy is 90 in all versions 
	Encore doesn't end right after a switch 
	GSCRBY Leech Seed damage doesn't constantly decrease 
	Berserk Gene's confusion is not perminant in Stadium Mode 
	Encore and Mirror Move are not Encorable in GSC 
	Terrain Selecter is now always enabled 
	Minor typos fixed here and there 
	Bugs in the Config Wizard are fixed 
	"Revert" menu item added to Team Builder 
	GSC Thief can now cause Flinching the same turn it steals King's Rock 
	GSC Unfreezing and Future Sight work 
	Restorative Present now checks the correct Poke's HP 
	GSC Rage now works accurately 
	Equal-speed randomization fixed (So THAT's how my Jolt was haxing all those Aeros!) 
	Deoxys F/L Pokedex number glitch fixed 
	Bide with 0 PP glitch fixed 
	GSCRBY Poison and Burns work off Max HP instead of Current HP now 
	Mirror Move glitch fixed 
	Protect fails if everyone else switches 
	Illegal Marill message fixed 
	Opening the Replay window doesn't screw the Battle window 
	Access Box now checks to make sure your team has changed before sending the TeamChange packet 
	Encore/Metronome weirdness fixed 
	Using Eruption/WaterSpout with 1 HP doesn't spew an RTE now 
	Pi/Pika/Raichu now knows Surf
	Note: Wish on Kirlia/Gardevoir is Japan-only, but is in because you can trade between US & Japan Ruby/Sapphire.
	Lovely Kiss has 10 PP 
	The Pokedex now remembers what mode you were in
	Details bit the Pokedex doesn't appear in RBY Mode
	Leech Seed miss message altered
	Leech Seed damage now rounds down
	Various spelling errors and other minor problems fixed
	GSCRBY Confusion duration error fixed
	SE/NVE Message moved BELOW the Damage message
	Compatibility checks fixed
	Present can't heal Ghost types now
	GSC Spite no longer freezes the game
	Foresight now works correctly
	GSC Unown letter changing now works correctly
	RBY Bugs:
		Disable now accounts for less than four moves 
		Critical Hit formula tweaked 
		Type Matching Chart has been touched up 
		Wrap Duration is now 2-5 instead of 3-6 
		RBY Teams in GSC Mode now have correct Special stats 
		OHKO Message comes before SE/NVE 
		Bite is now Counterable 
		Transform is now Mirror Movable 
		Rest/Recover/Softboiled returns "But it failed!" at full HP 
		Rest's 0/255/511 craziness also applies to Recover and Softboiled 
		Poison Sting's poisoning rate reduced to 20 
		RBY Rage's attack raising works just like Ruby/Sapp Rage's. Heh... cool 
		Switching during Bide is now possible 
		Countering OHKOs removed from RBY 
		Struggle does 50% recoil 
		Disable duration changed to 1 to 4 
		Zero Damage attacks now generate a miss 
		Par/Slp/Frz/etc doesn't reset Counter damage 
		Moves called by Move Callers are now Mirror Movable. 
		Psywave damage formula tweaked 
		Two-Turn moves now deplete a PP on the second turn instead of the first 
		Thrash and Petal Dance continue even if they miss 
		No move selection is allowed while Asleep or Frozen 
		Mirror Move now mirrors moves called by Move Callers 
		RBY multihit moves now stop after breaking a Substitute 
	Script Fixes/Additions:
		BattleOver fires for Unrated Forfeits and Disconnects 
		The $Message variable in BattleOver has a "*" after it for Unrated battles (eg, "WIN*" means Unrated Win) 
		ChallengeIssued works again 
		#IsIn gives a position number 
		/me fires -ChatMessage 
		#GetCompat's numbers have been updated 
		PAs are cleared when a player signs on 
		New Functions:
			#Round(#Val, #Places) 
			#PokeNum($Pokemon) 
			#MoveNum($Move) 
			$Move(#N) 
			$Item(#N) 
			#GetPokeLevel(#PNum, #PokeNum) 
			#GetPokeItem(#PNum, #PokeNum)
0.9.2G (Test Release)
	Breed Check tweaks
	Own Tempo message
	Spectator HP Bars
	RBY Teams in GSC Mode
	Fly/Thrash Problem
	Perish Song Final Pokes
	Protect/SpitUp
	Leech Seed
	Hydro Pump/Hyper Beam Accuracy
	RBY Team Builder RTE
	Minor Script fixes
	Tempban/Permban glitch
	Other minor fixes
	RBY Bugs:
		Ghost/Wrap Duration
		OHKO Moves NVE Message
		Damage Rounding
		Rest/Toxic
0.9.2F (Test Release)
	Modal/Non-Modal IM Window RTE
	Control Window RTE
	Box Poke Item Bug
	File>New Bug
	Disable Watch Chat DC
	Imprison
	Transform Gender
	Transform/Sketch
	Encore/Torment
	PoisonPoint/Steel Immunity
	Protect/Rollout
	Bide/Substitute
	Protect/Switching
	RuSa Low Kick
	Clamp message
	Present RTE
	RuSa Psywave damage
	Spectator RTE
	Baby Pokemon egg moves
	Various database fixes (mostly L/F Tutors)
	RBY Bugs:
		Acc/Evade Tooltip
		Item removed from Tooltip
		Wrap/Bind vs Ghost
		Counter Acc.
		Explosion/Substitute
		Dizzy Punch Effect
		Hi Jump Kick Accuracy/CrashDamage
		Struggle Type
		Roar/Whirlwind Priority
		Leech Seed
		Haze/Toxic
		RBY Wrap
		Wrap with 1 PP
0.9.2E (Test Release)
	Struggle Bug (99.6% sure)
	Battle ChatBox RTE 52
	Single-Space Name RTE
	Spit Up RTE
	Stadium Mode
	Growl/Substitute
	AilmentTraits/Substitute
	Trace
	Twineedle Poison
	Missing Self-Affecting moves
	Tail Glow
	Swagger/Flatter Substitute fixes
	Roar/Nature Cure
	Mean Look/Baton Pass
	Future Sight/Substitute
	Transient Grudge Effect
	Brick Break
	Transform Level copy
	Refresh Par/Brn Penalty
	Imprison
	Struggle/Rock Head
	Solar/Hyper Beam on final Doubles poke
	Heal Bell on final Doubles poke
	Focus Punch on final Doubles poke
	Non-RuSa Thrash
	Non-RuSa Counter/OHKOs
	RBY Rest
	RBY Disable
	RBY Recharge on Substitute break
	Tutor Moves updated
	Shedinja Fixes
	Dizzy Punch Jynx fixed
	Deoxys Box Fix
	Mist ToolTip text
	Helping Hand tooltip
	Multihitter damage message
	LeechSeed/LiquidOoze message
	Blank Watch Leave message
	Skip Delay inconsistancy
	Watch Chat Ignore inconsistancy
	Color Names fix
	Compatibility fix for Full Advance
	Special move descriptions display in all modes
	Team Builder reset on mode change
0.9.2D (Test Release)
	Struggle Bug (Probably - if it is, servers will be available in 0.9.2E)
	Shell Bell message
	Syncronized Par/Brn Stat Drop
	Battle Results Toggle
	Post-Faint LightningRod
	File>New bug
	Challenge Cup Items
	Spikes/Roar
	Yawn/SleepClause
	Level Balance
	Watch Battle graphics
	Freeze clause
	Delay Disable
	Numerous typos
	Database fixes for moves
	Installer copies an extra file
	Sleep Talk GSC PP Depletion
	RBY Fixes:
		Counter
		LS/Reflect
		Bite 10% Flinch Rate
0.9.2C (Test Release)
	Struggle (Partially - we'll need debug logs for the times it comes up when you have PP left)
	File not Found error (Probably - it looks like the installer messed up if you installed to the same folder, which I did say not to do...)
	Critical Hits
	Challenge Cup
	LightningRod
	Baton Pass
	Hidden Power
	Solar Beam
	Syncronize/Poison
	PP Display
	Collesium Images
	Unrated Battle
	Team Builder freeze
	PM RTE
	Single Battle Free Move
	Status when dead
	Rain Dish
	Reflect/Light Screen
	Encore
	über markers on hidden pokes
	Destiny Bond
	Pokedex images
	PM RTE Index Out of Bounds
	NightShade/Shedinja
	QuickClaw
	Team Power RTE
	Box-Of-White-Dots syndrome
	SandAttack/Sandstorm Dex data
	Sandstorm Damage
	Helping Hand
	Moonlight
	Yawn
	Toxic Message
	Post-Faint LeechSeed/Wish
	Water Spout
	Beat Up
	Pursuit
	Pokedex Gender Display
	Blank-Name Watcher Disconnect
	Mirror Move
	Volt/Water Absorb Text
	Image fixes (including Colosseum)
	RBY Bugs:
		Pyschic's Spcl. Fall rate
		Low Kick
		Transform
		Bind/Wrap/FireSpin/Clamp
0.9.2B (Test Release)
	Runtime 87
	Rest
	Randomly targeting moves
	Moves that were showing up as "not working" in the Team Builder
0.9.2A (Test Release)
	First Advance test release, lots of new stuff
0.9.2
	Fixed a major bug with Away status
0.9.1
	Some bugs found with 0.9.0 fixed (especially networking)
	Evolution data finished & corrected in Pokedex
0.9.0
	Minor fixes from 0.9.H
0.9.H (Release Candidate 1)
	Fixed bugs with 0.9.G
	Added server verison check
0.9.G (Test Release)
	Fixed bugs with 0.9.F
0.9.F (Test Release)
	Fixed bugs with .E
	Included correct database (oops)
	Improved Tooltips
	Challenge Cup mode (Random Pokemon)
0.9.E (Test Release)
	Fixed bugs with .D
0.9.D (Test Release)
	Fixed bugs with .C
	Autosave/Prompt to Save expanded
	Replays openable in Explorer
	Compatibility shown in Challenge window
0.9.C (Test Release)
	Fixed bugs with .B
	Optimized network system
	Refined Advance graphics
	Added MoveDex
	Expanded Translator
0.9.B (Test release)
	Fixed bugs with .A
	Added Advance graphics
0.9.A (Test release)
	Major changes
0.8.43
	Fixed problem with language files 
	Included correct database (PokéCenter moves, etc.)
0.8.42
	Variable damage moves fixed (Magnitude, Present, Return, etc.) 
	Fixed a crash when you double-click your own name on the server 
	Even more Substitute fixes
	Reflect/Light Screen fixed
0.8.41
	/me, /ignore, /unignore added 
	Speed checking added 
	Server tray icon added 
	Server can limit the text in the window (fixes a memory leak) 
	Server can disable queuing (EXPERIMENTAL!!!) 
	Pokédex fixed 
	Struggle fixed 
	Transform fixed 
	Spite fixed 
	Rollout fixed 
	Stomp/Minimize fixed 
	Quick Claw priority fixed 
	Future Sight fixed 
	Ghost-types' Curse fixed 
	Recoil damage moves fixed 
	PP Up rule fixed 
	Berserk Gene fixed 
	Critical Hit calculation modified 
	Confusion fixed 
	More substitute fixes 
	Status conditions fixed 
	Hidden Power fixed to work with Counter 
	Destiny Bond fixed 
	Text scrolling fixed


Credits
---------------------

Lead Programmer
* TV's Ian
Additional programming
* MasamuneXGP
* Don't Run With Scizors
* Jshadias
* mleo2003
Database Entry/Additional Help
* ShadowHawk
* Evil Gibson
* Tales9
* Mana Lugia
* Nautilator

Pokemon is (C) Nintendo, Game Freak, CREATURES, Inc., et al.  NetBattle was created
using publicly available information and experimentation with the Pokemon games, due to a
lack of Internet-enabled Pokemon games.