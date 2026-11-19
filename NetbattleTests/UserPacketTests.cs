using Microsoft.VisualStudio.TestTools.UnitTesting;
using Netbattle.Network;

namespace NetbattleTests {
    [TestClass]
    public class UserPacketTests {
        [TestMethod]
        public void EncodeSavedPassword_EmptyIs15Nulls() { // -- VB: String$(15, vbNullChar)
            CollectionAssert.AreEqual(new byte[15], UserPacket.EncodeSavedPassword(""));
            CollectionAssert.AreEqual(new byte[15], UserPacket.EncodeSavedPassword(null));
        }

        [TestMethod]
        public void EncodeSavedPassword_HexPairsBecome16Bytes() { // -- VB: For X = 1 To 32 Step 2: Chr$(Val("&H" & Mid(Temp, X, 2)))
            // -- MD5("password")
            byte[] result = UserPacket.EncodeSavedPassword("5F4DCC3B5AA765D61D8327DEB882CF99");

            CollectionAssert.AreEqual(new byte[] {
                0x5F, 0x4D, 0xCC, 0x3B, 0x5A, 0xA7, 0x65, 0xD6,
                0x1D, 0x83, 0x27, 0xDE, 0xB8, 0x82, 0xCF, 0x99
            }, result);
        }

        [TestMethod]
        public void EncodeSavedPassword_LowercaseHexWorks() {
            CollectionAssert.AreEqual(
                UserPacket.EncodeSavedPassword("5F4DCC3B5AA765D61D8327DEB882CF99"),
                UserPacket.EncodeSavedPassword("5f4dcc3b5aa765d61d8327deb882cf99"));
        }
    }
}
