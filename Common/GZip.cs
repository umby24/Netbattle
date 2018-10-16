using System;
using System.IO;
using System.IO.Compression;

namespace Netbattle.Common {
    public static class GZip {
        public static  byte[] Decompress(byte[] compressed) {
            try {
                byte[] output;

                using (var ms = new MemoryStream()) {
                    // -- Output stream..
                    using (var zip = new GZipStream(new MemoryStream(compressed), CompressionMode.Decompress)) {
                        var buffer = new byte[1024];
                        while (true) {
                            int bytesRead = zip.Read(buffer, 0, 1024);
                            if (bytesRead == 0) break;
                            ms.Write(buffer, 0, bytesRead);
                        }
                    }
                    output = ms.ToArray();
                }

                return output;

            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        public static byte[] Decompress2(byte[] compressed) {
            try {
                byte[] output;
                
                using (var ms = new MemoryStream()) {
                    // -- Output stream..
                    using (var zip = new DeflateStream(new MemoryStream(compressed), CompressionMode.Decompress)) {
                        var buffer = new byte[1024];
                        while (true) {
                            int bytesRead = zip.Read(buffer, 0, 1024);
                            if (bytesRead == 0) break;
                            ms.Write(buffer, 0, bytesRead);
                        }
                    }
                    output = ms.ToArray();
                }

                return output;

            }
            catch (Exception e) {
                Console.WriteLine(e.Message);
                return null;
            }
        }
    }
}
