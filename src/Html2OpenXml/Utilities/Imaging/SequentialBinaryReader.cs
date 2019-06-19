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
 * Inspiration from Metadata Extractor (Drew Noakes):
 * https://github.com/drewnoakes/metadata-extractor-dotnet
 */

using System.IO;

namespace HtmlToOpenXml
{
    /// <summary>
    /// Reads primitive data types as binary values with endianness support.
    /// </summary>
    sealed class SequentialBinaryReader : BinaryReader
    {
        public bool IsBigEndian { get; set; }


        public SequentialBinaryReader(Stream input) : base(input)
        {
        }

        /// <summary>
        /// Skips forward in the sequence.
        /// </summary>
        public void Skip (int count)
        {
            if (BaseStream.CanSeek) BaseStream.Seek(count, SeekOrigin.Current);
            else this.ReadBytes(count);
        }

        public override ushort ReadUInt16()
        {
            if (this.IsBigEndian)
                return  (ushort) (ReadByte() << 8 | ReadByte());
            return (ushort) (ReadByte() | ReadByte() << 8);
        }

        public override short ReadInt16()
        {
            if (this.IsBigEndian)
                return  (short) (ReadByte() << 8 | ReadByte());
            return (short) (ReadByte() | ReadByte() << 8);
        }

        public override int ReadInt32()
        {
            if (this.IsBigEndian)
                return ReadByte() << 24 | ReadByte() << 16 | ReadByte() << 8  | ReadByte();
            return ReadByte() | ReadByte() <<  8 | ReadByte() << 16 | ReadByte() << 24;
        }

        public override uint ReadUInt32()
        {
            if (this.IsBigEndian)
                return (uint) (ReadByte() << 24 | ReadByte() << 16 | ReadByte() << 8  | ReadByte());
            return (uint) (ReadByte() | ReadByte() <<  8 | ReadByte() << 16 | ReadByte() << 24);
        }
    }
}
