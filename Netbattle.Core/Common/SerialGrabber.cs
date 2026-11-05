using System.Buffers.Binary;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace Netbattle.Core.Common {
    /// <summary>
    /// Port of clsSerialGrabber.cls: reads the first hard drive's serial number, which NbMethods.GetSerialNumber hashes
    /// into the Station ID. Byte handling follows the VB exactly, quirks included, so a machine gets the same SID it got
    /// from the original client. The drive reads are Windows-only; the byte parsing is not. The Win9x VxD path
    /// (ReadUsingVxDinWin9X) isn't ported: .NET doesn't run there.
    /// </summary>
    public class SerialGrabber {
        #region Win32
        private const uint GenericRead = 0x80000000;
        private const uint GenericWrite = 0x40000000;
        private const uint FileShareRead = 0x1;
        private const uint FileShareWrite = 0x2;
        private const uint OpenExisting = 3;

        private const uint DfpGetVersion = 0x74080;
        private const uint DfpReceiveDriveData = 0x7C088;
        private const uint IoctlStorageQueryProperty = 0x2D1400;
        private const uint IoctlScsiMiniport = 0x4D008;
        private const int IoctlScsiMiniportIdentify = (0x1B * 0x10000) + 0x501; // -- FILE_DEVICE_SCSI * &H10000 + &H501

        private const byte IdeAtapiIdentify = 0xA1;
        private const byte IdeAtaIdentify = 0xEC;
        private const int IdentifyBufferSize = 512;
        private const int OutputDataSize = IdentifyBufferSize + 16;

        // -- Sizes as VB's Len() gives them, since the VB passes those to DeviceIoControl.
        private const int SendCmdInParamsLength = 36; // -- 32 bytes of fields plus the 4-byte bBuffer() array pointer.
        private const int SendCmdOutParamsLength = 529; // -- 4 + DRIVERSTATUS (12) + bBuffer(512) (513)
        private const int SrbIoControlLength = 28;

        // -- Offsets into the structures.
        private const int SendCmdOutBuffer = 16; // -- SENDCMDOUTPARAMS.bBuffer
        private const int IdSectorSerial = 20; // -- IDSECTOR.sSerialNumber, 20 bytes
        private const int SerialLength = 20;
        private const int StorageDescriptorSerialOffset = 24; // -- STORAGE_DEVICE_DESCRIPTOR.SerialNumberOffset

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(string fileName, uint desiredAccess, uint shareMode, IntPtr securityAttributes,
            uint creationDisposition, uint flagsAndAttributes, IntPtr templateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeviceIoControl(SafeFileHandle device, uint ioControlCode, byte[]? inBuffer, int inBufferSize,
            byte[] outBuffer, int outBufferSize, out int bytesReturned, IntPtr overlapped);
        #endregion

        /// <summary>
        /// Tries each method in the VB's order (admin IDENTIFY, SCSI miniport, then the no-rights storage query) and returns
        /// the first serial found, or "" when none works. A method that throws counts as finding nothing, like the VB's
        /// On Error Resume Next.
        /// </summary>
        [SupportedOSPlatform("windows")]
        public string GetSerial() {
            string serial = TryRead(() => ReadUsingXPAdminRights());

            if (serial.Length == 0)
                serial = TryRead(ReadUsingSCSIBackdoor);

            if (serial.Length == 0)
                serial = TryRead(() => ReadUsingXPZeroRights());

            return serial;
        }

        private static string TryRead(Func<string> reader) {
            try {
                return reader();
            }
            catch (Exception e) {
                Logger.Log(LogType.Debug, $"[SerialGrabber] {e.Message}");
                return "";
            }
        }

        /// <summary>
        /// Sends an ATA (or ATAPI) IDENTIFY through SMART_RCV_DRIVE_DATA. Needs read/write access to the drive, so usually
        /// administrator rights.
        /// </summary>
        [SupportedOSPlatform("windows")]
        public string ReadUsingXPAdminRights(int deviceId = 0) {
            using SafeFileHandle drive = OpenDevice($@"\\.\PhysicalDrive{deviceId}", GenericRead | GenericWrite);
            if (drive.IsInvalid)
                return "";

            // -- GETVERSIONOUTPARAMS. The result is ignored like the VB: on failure the device map stays 0 and ATA is used.
            var versionParams = new byte[24];
            DeviceIoControl(drive, DfpGetVersion, null, 0, versionParams, versionParams.Length, out _, IntPtr.Zero);
            byte devMap = versionParams[3]; // -- bIDEDeviceMap

            byte idCmd = ((devMap >> deviceId) & 16) != 0 && devMap > 0 ? IdeAtapiIdentify : IdeAtaIdentify;

            byte[] output = new byte[OutputDataSize];
            if (!DoIdentify(drive, output, idCmd, (byte)deviceId))
                return "";

            return SerialFromIdSector(output, SendCmdOutBuffer);
        }

        /// <summary>
        /// Reads the serial from the storage device descriptor. Opens the drive with no access rights, so it works without
        /// administrator rights.
        /// </summary>
        [SupportedOSPlatform("windows")]
        public string ReadUsingXPZeroRights(int deviceId = 0) {
            using SafeFileHandle drive = OpenDevice($@"\\.\PhysicalDrive{deviceId}", 0);
            if (drive.IsInvalid)
                return "";

            // -- StoragePropertyQuery: PropertyId = StorageDeviceProperty (0), QueryType = PropertyStandardQuery (0). The VB's
            // -- Len() counts the AdditionalParameters() pointer, so 12 bytes.
            var query = new byte[12];
            var buffer = new byte[10000];

            if (!DeviceIoControl(drive, IoctlStorageQueryProperty, query, query.Length, buffer, buffer.Length, out int bytesReturned, IntPtr.Zero))
                return "";

            return SerialFromStorageDescriptor(buffer, bytesReturned);
        }

        /// <summary>
        /// Sends an ATA IDENTIFY through each SCSI controller's miniport driver (\\.\Scsi0: to \\.\Scsi17:, drives 0 and 1).
        /// </summary>
        [SupportedOSPlatform("windows")]
        public string ReadUsingSCSIBackdoor() {
            // -- SRB_IO_CONTROL, then SENDCMDINPARAMS going in and SENDCMDOUTPARAMS coming back.
            var buffer = new byte[SrbIoControlLength + SendCmdOutParamsLength + IdentifyBufferSize];

            for (var controller = 0; controller <= 17; controller++) {
                using SafeFileHandle device = OpenDevice($@"\\.\Scsi{controller}:", GenericRead | GenericWrite);
                if (device.IsInvalid)
                    continue;

                for (byte driveNum = 0; driveNum <= 1; driveNum++) {
                    Span<byte> span = buffer;
                    BinaryPrimitives.WriteInt32LittleEndian(span[0..], SrbIoControlLength); // -- HeaderLength
                    Encoding.ASCII.GetBytes("SCSIDISK").CopyTo(span[4..]); // -- Signature
                    BinaryPrimitives.WriteInt32LittleEndian(span[12..], 10000); // -- Timeout
                    BinaryPrimitives.WriteInt32LittleEndian(span[16..], IoctlScsiMiniportIdentify); // -- ControlCode
                    BinaryPrimitives.WriteInt32LittleEndian(span[24..], SendCmdOutParamsLength + IdentifyBufferSize); // -- Length

                    // -- Only bCommandReg and bDriveNumber are set in SENDCMDINPARAMS here. The VB never clears the rest of
                    // -- the buffer between attempts, and neither does this.
                    span[SrbIoControlLength + 4 + 6] = IdeAtaIdentify; // -- irDriveRegs.bCommandReg
                    span[SrbIoControlLength + 12] = driveNum; // -- bDriveNumber

                    if (!DeviceIoControl(device, IoctlScsiMiniport, buffer, SrbIoControlLength + SendCmdInParamsLength - 1,
                            buffer, buffer.Length, out _, IntPtr.Zero))
                        continue;

                    if (BinaryPrimitives.ReadInt32LittleEndian(span[SrbIoControlLength..]) <= 0) // -- SENDCMDOUTPARAMS.cBufferSize
                        continue;

                    string serial = SerialFromIdSector(buffer, SrbIoControlLength + SendCmdOutBuffer);
                    if (serial.Length != 0)
                        return serial;
                }
            }

            return "";
        }

        private static SafeFileHandle OpenDevice(string name, uint access) {
            return CreateFile(name, access, FileShareRead | FileShareWrite, IntPtr.Zero, OpenExisting, 0, IntPtr.Zero);
        }

        private static bool DoIdentify(SafeFileHandle drive, byte[] output, byte idCmd, byte driveNum) {
            // -- SENDCMDINPARAMS
            var input = new byte[SendCmdInParamsLength];
            BinaryPrimitives.WriteInt32LittleEndian(input, IdentifyBufferSize); // -- cBufferSize
            input[4] = 0; // -- bFeaturesReg
            input[5] = 1; // -- bSectorCountReg
            input[6] = 1; // -- bSectorNumberReg
            input[7] = 0; // -- bCylLowReg
            input[8] = 0; // -- bCylHighReg
            input[9] = (byte)(0xA0 | ((driveNum & 1) << 4)); // -- bDriveHeadReg
            input[10] = idCmd; // -- bCommandReg
            input[12] = driveNum; // -- bDriveNumber

            return DeviceIoControl(drive, DfpReceiveDriveData, input, SendCmdInParamsLength - 4, output, OutputDataSize, out _, IntPtr.Zero);
        }

        /// <summary>
        /// Pulls the serial out of an IDENTIFY sector starting at <paramref name="offset"/>: swaps each byte pair (ATA strings
        /// are big-endian words) and cuts at the first NUL. Padding spaces are kept, like the VB.
        /// </summary>
        public static string SerialFromIdSector(byte[] data, int offset) {
            byte[] serial = data.AsSpan(offset + IdSectorSerial, SerialLength).ToArray();
            return CutAtNull(Encoding.Latin1.GetString(SwapBytes(serial)));
        }

        /// <summary>
        /// Pulls the serial out of a STORAGE_DEVICE_DESCRIPTOR buffer and decodes it with FlipAndCodeBytes. Returns "" when the
        /// serial is under 2 bytes long, or when its length isn't a multiple of 4 (the VB reads past the end of its array
        /// there, and the error makes that read return nothing).
        /// </summary>
        public static string SerialFromStorageDescriptor(byte[] buffer, int bytesReturned) {
            int serialOffset = BinaryPrimitives.ReadInt32LittleEndian(buffer.AsSpan(StorageDescriptorSerialOffset));
            if (serialOffset < 0)
                return "";

            // -- For X = SerialNumberOffset To cbBytesReturned: stops on the first NUL, or ends one past cbBytesReturned.
            int end = serialOffset;
            while (end <= bytesReturned && buffer[end] != 0)
                end++;

            if (end <= serialOffset + 1)
                return "";

            byte[] serial = buffer[serialOffset..end];
            if (serial.Length % 4 != 0)
                return "";

            return Encoding.Latin1.GetString(FlipAndCodeBytes(serial));
        }

        /// <summary>
        /// Swaps every pair of bytes in place and returns the same array.
        /// </summary>
        public static byte[] SwapBytes(byte[] bytes) {
            for (var i = 0; i + 1 < bytes.Length; i += 2)
                (bytes[i], bytes[i + 1]) = (bytes[i + 1], bytes[i]);

            return bytes;
        }

        /// <summary>
        /// Decodes the hex-text serial that XP returns from the storage query: each 4 characters are two hex bytes, taken
        /// second pair first. Only 0-9 and lowercase a-f count as digits (anything else adds 0), and zero bytes are dropped.
        /// The caller makes sure the length is a multiple of 4.
        /// </summary>
        public static byte[] FlipAndCodeBytes(byte[] bytes) {
            var flipped = new List<byte>();

            for (var i = 0; i < bytes.Length; i += 4) {
                for (var j = 1; j >= 0; j--) {
                    var sum = 0;
                    for (var k = 0; k <= 1; k++)
                        sum = sum * 16 + HexDigit(bytes[i + j * 2 + k]);

                    if (sum > 0)
                        flipped.Add((byte)sum);
                }
            }

            return flipped.ToArray();
        }

        private static int HexDigit(byte b) {
            if (b >= '0' && b <= '9')
                return b - '0';

            if (b >= 'a' && b <= 'f')
                return b - 'a' + 10;

            return 0;
        }

        private static string CutAtNull(string value) {
            int nul = value.IndexOf('\0');
            return nul < 0 ? value : value[..nul];
        }
    }
}
