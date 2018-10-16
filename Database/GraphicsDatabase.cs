using System;
using System.IO;
using System.Text;
using Netbattle.Common;

namespace Netbattle.Database {
    /// <summary>
    /// Database for the loading and managing of all pokemon sprites
    /// </summary>
    public class GraphicsDatabase {
        public static GraphicsData GraphicsMap;

        public static void Load() {
            GraphicsMap = new GraphicsData();

            using (var fs = new FileStream("graphics.bin", FileMode.Open)) {
                using (var bs = new BinaryReader(fs)) {
                    var gfxBytes = new byte[4];
                    bs.Read(gfxBytes, 0, 4);

                    // -- shorts? (There are two, we only need 1.. reading it from the stream using .ReadInt16() messes it up massively.
                    int compressedLength = gfxBytes[0] * 256 + gfxBytes[1]; // -- 'X' in the original code

                    bs.ReadInt16(); // -- Strip GZip header.

                    // -- Read and decompress sprite names
                    var temp = new byte[compressedLength - 2];
                    bs.Read(temp, 0, compressedLength - 2);
                    temp = GZip.Decompress2(temp);
                    // -- Split into individual values.
                    var gfxHeader = Encoding.UTF8.GetString(temp).Split('|');
                    GraphicsMap.Titles = gfxHeader;
                    compressedLength = gfxHeader.Length;
                    int uncompressedLength = (compressedLength * 12) / 8; // -- The number of sets of 12 we have.

                    // -- Read in the image lengths.
                    temp = new byte[uncompressedLength];
                    bs.Read(temp, 0, uncompressedLength - 1);

                    string imageLengths = NbMethods.BytesToBinary(temp);// -- expand each character into a bit string..
                    GraphicsMap.ByteCount = new long[compressedLength + 1];
                    GraphicsMap.ByteStart = new long[compressedLength + 1];
                    GraphicsMap.InFile = new byte[compressedLength + 1];
                    var z = 0;

                    // -- this grabs the bits, 12 at a time, converts those to an Int, and that's the number of bytes this GIF is.
                    for (var y = 0; y < compressedLength; y++) {
                        // -- Get the value off the top of the string
                        string build = "0000" + imageLengths.Substring(0, 12);
                        // -- update the string for the next usage.
                        imageLengths = imageLengths.Substring(12);
                        // -- Convert to bytes
                        byte[] bytesAgain = NbMethods.BinaryToBytes(build);
                        // -- Convert to short/int.
                        var size = 0;
                        size |= bytesAgain[0] << 8;
                        size |= bytesAgain[1];
                        // -- Add value to db.
                        GraphicsMap.ByteStart[y] = z;
                        GraphicsMap.ByteCount[y] = size;
                        z += size;
                        GraphicsMap.InFile[y] = 1;
                    }

                    // -- The whole rest of the file is the actual GZIP compressed data now.
                    long bytesToRead = bs.BaseStream.Length - bs.BaseStream.Position;
                    var finalData = new byte[bytesToRead - 2];
                    bs.ReadInt16();
                    bs.Read(finalData, 0, (int)(bytesToRead - 2));
                    File.WriteAllBytes("temp.bin", finalData);

                    byte[] decompressedFinal = GZip.Decompress2(finalData);
                    File.WriteAllBytes("gfxtmp.bin", decompressedFinal); // -- TODO: Just keep this in memory? OG NB just dumps it to a temp file and reads from it when needed.
                }
            }
        }

        public static byte[] GetSprite(int index) {
            long spriteLocation = GraphicsMap.ByteStart[index];
            long spriteLength = GraphicsMap.ByteCount[index];
            var resultBytes = new byte[spriteLength];

            using (var fs = new FileStream("gfxtmp.bin", FileMode.Open)) {
                using (var bs = new BinaryReader(fs)) {
                    fs.Seek(spriteLocation, SeekOrigin.Begin);
                    bs.Read(resultBytes, 0, (int)spriteLength);
                }
            }

            return resultBytes;
        }

        public static byte[] GetSprite(string spriteName) {
            var index = Array.IndexOf(GraphicsMap.Titles, spriteName) - 1;
            return GetSprite(index);
        }

        public static byte[] GetSprite(Pokemon poke, GraphicsMode graphics, bool backView, byte weather = 0) {
            if (poke.No == 201) {
                return GetUnknownSprite(poke, graphics, backView);
            }

            if (poke.No == 351) {
                return GetCastformSprite(poke, graphics, backView, weather);
            }

            var imageName = poke.No.ToString().PadLeft(3, '0');
            bool canShiny = true;

            switch (graphics) {
                case GraphicsMode.nbGFXRB:
                    imageName += "rb";
                    canShiny = false;
                    break;
                case GraphicsMode.nbGFXGrn:
                    imageName += "rg";
                    canShiny = false;
                    break;
                case GraphicsMode.nbGFXYlo:
                    imageName += "y";
                    canShiny = false;
                    break;
                case GraphicsMode.nbGFXGld:
                    imageName += "g";
                    break;
                case GraphicsMode.nbGFXSil:
                    imageName += "s";
                    break;
                case GraphicsMode.nbGFXRS:
                    imageName += "rs";
                    break;
                case GraphicsMode.nbGFXLF:
                    imageName += "fl";
                    break;
                case GraphicsMode.nbGFXCol:
                    imageName += "c";
                    break;
                case GraphicsMode.nbGFXEme:
                    imageName += "e";
                    break;
                case GraphicsMode.nbGFXSml:
                    imageName += "_1";
                    canShiny = false;
                    break;
            }

            if (backView && graphics != GraphicsMode.nbGFXSml)
                imageName += "b";

            if (canShiny && poke.Shiny)
                imageName += "s";

            return GetSprite(imageName);
        }

        private static byte[] GetCastformSprite(Pokemon poke, GraphicsMode graphics, bool backview, byte Weather) {
            var spriteName = "";

            switch (graphics) {
                case GraphicsMode.nbGFXSml:
                    spriteName = "351_1";
                    break;
                case GraphicsMode.nbGFXCol:
                    if (Weather == 0 || Weather == 3) {
                        spriteName = "351c";
                        if (poke.Shiny)
                            spriteName += "s";
                        break;
                    }

                    goto default;
                default:
                    spriteName = "351rs";
                    if (backview)
                        spriteName += "b";
                    if ((Weather == 0 || Weather == 3) && poke.Shiny) {
                        spriteName += "s";
                    } else if (Weather != 0 && Weather != 3) {
                        spriteName += Weather.ToString();
                    }

                    break;
            }

            return GetSprite(spriteName);
        }

        private static byte[] GetUnknownSprite(Pokemon poke, GraphicsMode graphics, bool backView) {
            var spriteName = "";

            switch (graphics) {
                case GraphicsMode.nbGFXSml:
                    var temp = (char) (poke.UnownLetter + 97);
                    var build = "";
                    if (temp == '{')
                        build = "ep";
                    if (temp == '|')
                        build = "qw";
                    spriteName = $"201{build}_1";
                    break;
                case GraphicsMode.nbGFXSil:
                case GraphicsMode.nbGFXGld:
                    spriteName = $"201{poke.UnownLetter + 1}00";
                    if (backView)
                        spriteName += "b";

                    break;
                case GraphicsMode.nbGFXCol:
                    spriteName = $"201{poke.UnownLetter + 1}00c";
                    break;
                default:
                    spriteName = $"unown{poke.UnownLetter + 1}00";
                    if (backView)
                        spriteName += "b";
                    break;
            }

            if (poke.Shiny)
                spriteName += "s";

            return GetSprite(spriteName);
        }
    }
}
