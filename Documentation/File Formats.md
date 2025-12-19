# PNB Files
PNB Files are used to store the user's current team and user information like win message, lose message, selected icon.


# CDB File
Compressed DataBase file
This file format is fairly straightforward.

It's the uncompressed length of the file as an ASCII String, terminated by CRLF, followed by a GZIP compressed CSV File.

You can view the C# Classes i've created to parse out each of the individual databases, Move, Types, and Pokemon.