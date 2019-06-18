/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 *
 * Original source code from Andy Wilson: http://www.codeproject.com/KB/cs/ReadingImageHeaders.aspx
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace HtmlToOpenXml
{
    /// <summary>
    /// Taken from http://stackoverflow.com/questions/111345/getting-image-dimensions-without-reading-the-entire-file/111349
    /// Minor improvements including supporting unsigned 16-bit integers when decoding Jfif and added logic
    /// to load the image using new Bitmap if reading the headers fails
    /// </summary>
    public static class ImageHeader
    {
        // https://en.wikipedia.org/wiki/List_of_file_signatures

        enum FileType { Unrecognized, Bitmap, Gif, Png, Jpeg }

        private static readonly byte[] pngSignatureBytes = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

        private static Dictionary<byte[], FileType> imageFormatDecoders = new Dictionary<byte[], FileType>()
        {
            { new byte[] { 0x42, 0x4D }, FileType.Bitmap },
            { Encoding.UTF8.GetBytes("GIF87a"), FileType.Gif },
            { Encoding.UTF8.GetBytes("GIF89a"), FileType.Gif }, // animated gif
            { pngSignatureBytes, FileType.Png },
            { new byte[] { 0xff, 0xd8 }, FileType.Jpeg }
        };

        private static readonly int MaxMagicBytesLength = imageFormatDecoders
            .Keys.OrderByDescending(x => x.Length).First().Length;

        /// <summary>
        /// Gets the dimensions of an image.
        /// </summary>
        /// <param name="stream">The content of the image.</param>
        /// <returns>The dimensions of the specified image.</returns>
        /// <exception cref="ArgumentException">The image was of an unrecognised format.</exception>
        public static Size GetDimensions(Stream stream)
        {
            using (SequentialBinaryReader reader = new SequentialBinaryReader(stream))
            {
                FileType type = DetectFileType (reader);
                stream.Seek(0L, SeekOrigin.Begin);
                switch (type)
                {
                    case FileType.Bitmap: return DecodeBitmap(reader);
                    case FileType.Gif: return DecodeGif(reader);
                    case FileType.Jpeg: return DecodeJfif(reader);
                    case FileType.Png: return DecodePng(reader);
                    default: return Size.Empty;
                }
            }
        }

        /// <summary>
        /// Resize an image keeping its aspect ratio.
        /// </summary>
        public static Size KeepAspectRatio(Size actualSize, Size preferredSize)
        {
            int width, height;

            // Resize by the highest difference ratio between constrained dimension and real one.
            bool forceResizeByWidth = preferredSize.Height <= 0 && preferredSize.Width > 0;
            bool forceResizeByHeight = preferredSize.Width <= 0 && preferredSize.Height > 0;
            if (forceResizeByWidth || (!forceResizeByHeight &&
                Math.Abs(preferredSize.Width - actualSize.Width) > Math.Abs(preferredSize.Height - actualSize.Height)))
            {
                width = preferredSize.Width;
                height = (int) (((float) actualSize.Height / actualSize.Width) * width);
            }
            else
            {
                height = preferredSize.Height;
                width = (int) (((float) actualSize.Width / actualSize.Height) * height);
            }

            return new Size(width, height);
        }

        /// <summary>
        /// Examines the a file's first bytes and estimates the file's type.
        /// </summary>
        private static FileType DetectFileType (SequentialBinaryReader reader)
        {
            byte[] magicBytes = new byte[MaxMagicBytesLength];
            for (int i = 0; i < MaxMagicBytesLength; i += 1)
            {
                magicBytes[i] = reader.ReadByte();
                foreach (var kvPair in imageFormatDecoders)
                {
                    if (StartsWith(magicBytes, kvPair.Key))
                    {
                        return kvPair.Value;
                    }
                }
            }

            return FileType.Unrecognized;
        }

        /// <summary>
        /// Determines whether the beginning of this byte array instance matches the specified byte array.
        /// </summary>
        /// <returns>Returns true if the first array starts with the bytes of the second array.</returns>
        private static bool StartsWith(byte[] thisBytes, byte[] thatBytes)
        {
            for (int i = 0; i < thatBytes.Length; i += 1)
            {
                if (thisBytes[i] != thatBytes[i])
                {
                    return false;
                }
            }

            return true;
        }

        private static Size DecodeBitmap(SequentialBinaryReader reader)
        {
            var magicNumber = reader.ReadUInt16();

            // skip past the rest of the file header
            reader.Skip(4 + 2 + 2 + 4);

            int headerSize = reader.ReadInt32();
            int width, height;

            // We expect the header size to be either 40 (BITMAPINFOHEADER) or 12 (BITMAPCOREHEADER)
            if (headerSize == 40)
            {
                // BITMAPINFOHEADER
                width = reader.ReadInt32();
                height = Math.Abs(reader.ReadInt32());
            }
            else if (headerSize == 12)
            {
                width = reader.ReadInt16();
                height = reader.ReadInt16();
            }
            else
            {
                // Unexpected DIB header size
                return Size.Empty;
            }

            return new Size(width, height);
        }

        private static Size DecodeGif(SequentialBinaryReader reader)
        {
            // 3 - signature: "GIF"
            // 3 - version: either "87a" or "89a"
            reader.Skip(6);

            int width = reader.ReadInt16();
            int height = reader.ReadInt16();
            return new Size(width, height);
        }

        private static Size DecodeJfif(SequentialBinaryReader reader)
        {
            reader.IsBigEndian = true;
            var magicNumber = reader.ReadUInt16(); // first two bytes should be JPEG magic number

            do
            {
                // Find next segment marker. Markers are zero or more 0xFF bytes, followed
                // by a 0xFF and then a byte not equal to 0x00 or 0xFF.
                byte segmentIdentifier = reader.ReadByte();
                byte segmentType = reader.ReadByte();

                // Read until we have a 0xFF byte followed by a byte that is not 0xFF or 0x00
                while (segmentIdentifier != 0xFF || segmentType == 0xFF || segmentType == 0)
                {
                    segmentIdentifier = segmentType;
                    segmentType = reader.ReadByte();
                }

                if (segmentType == 0xD9) // EOF?
                    return Size.Empty;

                // next 2-bytes are <segment-size>: [high-byte] [low-byte]
                var segmentLength = (int)reader.ReadUInt16();

                // segment length includes size bytes, so subtract two
                segmentLength -= 2;

                if (segmentType == 0xC0 || segmentType == 0xC2)
                {
                    reader.ReadByte(); // bits/sample, usually 8
                    int height = (int) reader.ReadUInt16();
                    int width = (int) reader.ReadUInt16();
                    return new Size(width, height);
                }
                else
                {
                    // skip this segment
                    reader.Skip(segmentLength);
                }
            }
            while (true);
        }

        private static Size DecodePng(SequentialBinaryReader reader)
        {
            reader.IsBigEndian = true;
            reader.ReadBytes(pngSignatureBytes.Length);
            reader.Skip(8);

            int width = reader.ReadInt32();
            int height = reader.ReadInt32();
            return new Size(width, height);
        }
    }
}
