using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Netbattle.Core.Common;

namespace NetbattleTests {
    [TestClass]
    public class SerialGrabberTests {
        [TestMethod]
        public void SwapBytes_SwapsEachPair() {
            CollectionAssert.AreEqual(new byte[] { 2, 1, 4, 3 }, SerialGrabber.SwapBytes(new byte[] { 1, 2, 3, 4 }));
        }

        [TestMethod]
        public void SerialFromIdSector_UnswapsAndKeepsPadding() {
            const string serial = "    S12345678ABCDEFG"; // -- 20 characters, space-padded like real drives
            byte[] sector = MakeIdentifyOutput(serial, 16);

            Assert.AreEqual(serial, SerialGrabber.SerialFromIdSector(sector, 16));
        }

        [TestMethod]
        public void SerialFromIdSector_CutsAtFirstNull() {
            byte[] sector = MakeIdentifyOutput("AB\0\0CDEFGHIJKLMNOPQR", 16);

            Assert.AreEqual("AB", SerialGrabber.SerialFromIdSector(sector, 16));
        }

        [TestMethod]
        public void SerialFromStorageDescriptor_DecodesXpHexSerial() {
            // -- "WD12" as XP gives it: hex text, with each pair of bytes swapped.
            (byte[] buffer, int returned) = MakeStorageDescriptor("44573231");

            Assert.AreEqual("WD12", SerialGrabber.SerialFromStorageDescriptor(buffer, returned));
        }

        [TestMethod]
        public void SerialFromStorageDescriptor_LengthNotMultipleOfFourGivesNothing() { // -- VB indexes past its array
            (byte[] buffer, int returned) = MakeStorageDescriptor("445732");

            Assert.AreEqual("", SerialGrabber.SerialFromStorageDescriptor(buffer, returned));
        }

        [TestMethod]
        public void SerialFromStorageDescriptor_NoSerialGivesNothing() {
            var buffer = new byte[100];
            BinaryPrimitives.WriteInt32LittleEndian(buffer, 40); // -- Version; SerialNumberOffset stays 0

            Assert.AreEqual("", SerialGrabber.SerialFromStorageDescriptor(buffer, 40));
        }

        [TestMethod]
        public void FlipAndCodeBytes_OnlyLowercaseHexCounts() { // -- Select Case 48-57, 97-102
            byte[] decoded = SerialGrabber.FlipAndCodeBytes(Encoding.ASCII.GetBytes("4A57"));

            Assert.AreEqual("W@", Encoding.Latin1.GetString(decoded)); // -- "4A" reads as 0x40
        }

        [TestMethod]
        public void FlipAndCodeBytes_DropsZeroBytes() {
            byte[] decoded = SerialGrabber.FlipAndCodeBytes(Encoding.ASCII.GetBytes("00570044"));

            Assert.AreEqual("WD", Encoding.Latin1.GetString(decoded));
        }

        [TestMethod]
        public void GetSerialNumber_IsFirst100BitsOfMd5() {
            byte[] hash = MD5.HashData(Encoding.ASCII.GetBytes("WD123ABC"));
            byte[] expected = hash.Take(12).Append((byte)(hash[12] & 0xF0)).ToArray();

            CollectionAssert.AreEqual(expected, NbMethods.GetSerialNumber("WD123ABC"));
        }

        [TestMethod]
        public void GetSerialNumber_KeepsOnlyZeroThroughUppercaseZ() { // -- Y >= 48 And Y <= 90
            CollectionAssert.AreEqual(NbMethods.GetSerialNumber("123"), NbMethods.GetSerialNumber("ab-12 3"));
        }

        [TestMethod]
        public void GetSerialNumber_DecompressesToTwentyOneCharacters() {
            Assert.AreEqual(21, NbMethods.DecompressSid(NbMethods.GetSerialNumber("WD123ABC")).Length);
        }

        private static byte[] MakeIdentifyOutput(string serial, int offset) {
            var output = new byte[offset + 512];
            byte[] bytes = SerialGrabber.SwapBytes(Encoding.Latin1.GetBytes(serial)); // -- ATA strings are big-endian words
            bytes.CopyTo(output, offset + 20);
            return output;
        }

        private static (byte[] buffer, int returned) MakeStorageDescriptor(string serial) {
            const int serialOffset = 40;
            var buffer = new byte[10000];
            BinaryPrimitives.WriteInt32LittleEndian(buffer, serialOffset); // -- Version
            BinaryPrimitives.WriteInt32LittleEndian(buffer.AsSpan(24), serialOffset); // -- SerialNumberOffset
            Encoding.ASCII.GetBytes(serial).CopyTo(buffer, serialOffset);
            return (buffer, serialOffset + serial.Length + 1);
        }
    }
}
