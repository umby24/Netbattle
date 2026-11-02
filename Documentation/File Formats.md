# PNB Files
PNB Files are used to store the user's current team and user information like win message, lose message, selected icon.


# CDB File
**C**ompressed **D**ata**B**ase file
This file format is fairly straightforward.

It's the uncompressed length of the file as an ASCII String, terminated by CRLF, followed by a GZIP compressed CSV File.

You can view the C# Classes i've created to parse out each of the individual databases, Move, Types, and Pokemon.

# Graphics.bin

This file contains all of the graphics for the pokemon in the game.

The first 2 bytes is a short representing the compressed length of the sprite name string.
[Station ID.md](Station%20ID.md)
The next 2 bytes is a short representing the uncompressed length of the sprite name string.

The next bytes are the sprite name string, which is GZIP deflate compressed. 
The string is a list of the sprite names separated by a pipe character "|".

Next is a binary blob, of length (# of sprite names * 12) / 8.

This blob is a set of 12-bit integer values concatenated together, each representing the UNCOMPRESSED length of the gif images contained next in the file.

The remaineder of this file is a gzip deflate compressed binary blob of gif images, concatenated together.

To process this, I'd reccomend reading the integer value blob bit by bit, extracting each 12-bit integer value, 
then reading that many bytes from the decompressed gif image blob to get each individual gif image.

In the original code, the entire uncompressed binary blob is written to file, 
and the offsets of each gif are caclulated and stored in a dictionary for later access.

Whenever a sprite is needed, the code looks up the offset of the sprite in the dictionary,
seeks to that position in the uncompressed binary blob, and reads the gif image from there.[Station ID.md](Station%20ID.md)