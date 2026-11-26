/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 *
 * Original source code from Andy Wilson: http://www.codeproject.com/KB/cs/ReadingImageHeaders.aspx
 * http://stackoverflow.com/questions/111345/getting-image-dimensions-without-reading-the-entire-file/111349
 * EMF Specifications: https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-emf/ae7e7437-cfe5-485e-84ea-c74b51b000be
 */
using System.Text;
using System.Xml.XPath;

namespace HtmlToOpenXml.IO;

/// <summary>
/// Utility methods used to inspect image streams and determine image characteristics
/// without loading the entire image.
/// 
/// These helpers are used internally by HtmlToOpenXml when processing images but may also
/// be useful when implementing custom image handling logic.
/// </summary>
static class ImageHeader
{
    // https://en.wikipedia.org/wiki/List_of_file_signatures

#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member
    public enum FileType { Unrecognized, Bitmap, Gif, Png, Jpeg, Emf, Xml }
#pragma warning restore CS1591

    private static readonly byte[] pngSignatureBytes = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    private static readonly Dictionary<byte[], FileType> imageFormatDecoders = new()
    {
        { new byte[] { 0x42, 0x4D }, FileType.Bitmap },
        { Encoding.UTF8.GetBytes("GIF87a"), FileType.Gif },
        { Encoding.UTF8.GetBytes("GIF89a"), FileType.Gif }, // animated gif
        { pngSignatureBytes, FileType.Png },
        { new byte[] { 0xff, 0xd8 }, FileType.Jpeg },
        { new byte[] { 0x1, 0, 0, 0 }, FileType.Emf },
        { Encoding.UTF8.GetBytes("<?xml "), FileType.Xml }, // Xml so potentially Svg
    };


    private static readonly int MaxMagicBytesLength = imageFormatDecoders
        .Keys.OrderByDescending(x => x.Length).First().Length;


    /// <summary>
    /// Read a image stream and try to detect its file type.
    /// </summary>
    /// <param name="stream">The readable image stream. Stream should support seeking.</param>
    /// <param name="type">The guess file type.</param>
    /// <returns>Returns true if the detection was successful.</returns>
    public static bool TryDetectFileType(Stream stream, out FileType type)
    {
        using var reader = new SequentialBinaryReader(stream, leaveOpen: true);
        type = DetectFileType(reader);
        if (type != FileType.Unrecognized)
        {
            stream.Seek(0L, SeekOrigin.Begin);
        }

        return type != FileType.Unrecognized;
    }

    /// <summary>
    /// Gets the dimensions of an image.
    /// </summary>
    /// <param name="stream">The readable image stream. Stream should support seeking.</param>
    /// <returns>The dimensions of the specified image.</returns>
    /// <exception cref="ArgumentException">The image is in an unrecognised format.</exception>
    public static Size GetDimensions(Stream stream)
    {
        using var reader = new SequentialBinaryReader(stream, leaveOpen: true);
        FileType type = DetectFileType(reader);

        if (type == FileType.Unrecognized)
            return Size.Empty;

        stream.Seek(0L, SeekOrigin.Begin);
        return type switch
        {
            FileType.Bitmap => DecodeBitmap(reader),
            FileType.Gif => DecodeGif(reader),
            FileType.Jpeg => DecodeJfif(reader),
            FileType.Png => DecodePng(reader),
            FileType.Emf => DecodeEmf(reader),
            FileType.Xml => DecodeXml(stream),
            _ => Size.Empty,
        };
    }

    /// <summary>
    /// Resize an image keeping its aspect ratio.
    /// </summary>
    public static (Unit Width, Unit Height) KeepAspectRatio(Size actualSize,
        long preferredWidthEmus, long preferredHeightEmus)
    {
        // Handle edge cases where dimensions are zero or negative
        if (actualSize.HasInvalidDimension)
            return (Unit.Empty, Unit.Empty);

        // Resize by the highest difference ratio between constrained dimension and real one.

        if (preferredWidthEmus > 0 && preferredHeightEmus == 0)
        {
            return (
                new Unit(UnitMetric.Emus, preferredWidthEmus),
                new Unit(UnitMetric.Emus,
                    Math.Round(
                        actualSize.Height * (double)preferredWidthEmus / actualSize.Width))
            );
        }

        if (preferredHeightEmus > 0 && preferredWidthEmus == 0)
        {
            return (
                new Unit(UnitMetric.Emus,
                    Math.Round(
                        actualSize.Width * (double)preferredHeightEmus / actualSize.Height)),
                new Unit(UnitMetric.Emus, preferredHeightEmus)
            );
        }

        return (
            new Unit(UnitMetric.Emus, preferredWidthEmus),
            new Unit(UnitMetric.Emus, preferredHeightEmus)
        );
    }

    /// <summary>
    /// Examines the first bytes of the file and estimates its image type if possible.
    /// </summary>
    /// <returns>Returns <see cref="FileType.Unrecognized"/> if not recognized.</returns>
    private static FileType DetectFileType (SequentialBinaryReader reader)
    {
        byte[] magicBytes = new byte[MaxMagicBytesLength];
        var availableBytes = reader.BaseStream.Length - reader.BaseStream.Position;
        // reasonably, we can assume that if we are at the end of the stream and we read the header,
        // the image content must be invalid or truncated.
        if (availableBytes < MaxMagicBytesLength)
            return FileType.Unrecognized;

        reader.Read(magicBytes, 0, MaxMagicBytesLength);

        var headerSpan = magicBytes.AsSpan();
        foreach (var kvPair in imageFormatDecoders)
        {
            // Determines whether the beginning of this array matches s known header.
            if (headerSpan.StartsWith(kvPair.Key))
            {
                return kvPair.Value;
            }
        }

        return FileType.Unrecognized;
    }

    private static Size DecodeBitmap(SequentialBinaryReader reader)
    {
        var _ = reader.ReadUInt16();

        // skip past the rest of the file header
        reader.Skip(4 + 2 + 2 + 4);

        int headerSize = reader.ReadInt32();
        uint width, height;

        // We expect the header size to be either 40 (BITMAPINFOHEADER) or 12 (BITMAPCOREHEADER)
        if (headerSize == 40)
        {
            // BITMAPINFOHEADER
            width = reader.ReadUInt32();
            height = (uint) Math.Abs(reader.ReadInt32());
        }
        else if (headerSize == 12)
        {
            width = reader.ReadUInt16();
            height = reader.ReadUInt16();
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

        var width = reader.ReadUInt16();
        var height = reader.ReadUInt16();
        return new Size(width, height);
    }

    private static Size DecodeJfif(SequentialBinaryReader reader)
    {
        reader.IsBigEndian = true;
        var _ = reader.ReadUInt16(); // first two bytes should be JPEG magic number

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
            int segmentLength = reader.ReadUInt16();

            // segment length includes size bytes, so subtract two
            segmentLength -= 2;

            if (segmentType == 0xC0 || segmentType == 0xC2)
            {
                reader.ReadByte(); // bits/sample, usually 8
                var height = reader.ReadUInt16();
                var width = reader.ReadUInt16();
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

        var width = reader.ReadUInt32();
        var height = reader.ReadUInt32();
        return new Size(width, height);
    }

    private static Size DecodeEmf(SequentialBinaryReader reader)
    {
        // EMR_HEADER: Type + Size + Bounds
        reader.Skip(4 + 4 + 16);

        // Frame: specify the rectangular inclusive-inclusive dimensions, 
        // in .01 millimeter units, of a rectangle that surrounds the image stored in the metafile.
        var left   = reader.ReadInt32();
        var top    = reader.ReadInt32();
        var right  = reader.ReadInt32();
        var bottom = reader.ReadInt32();

        // Skip other headers:
        // Signature (4) + Version (4) + Size (4) + Records (4) + Handles (2)
        // + nReserved (2) + nDescription (4) + offDescription (4) + PalEntries (4)
        reader.Skip(32);

        // Next 8 bytes specify the size of the reference device, in pixels
        var deviceSizeInPixelX = reader.ReadUInt32();
        var deviceSizeInPixelY = reader.ReadUInt32();

        // Then 8 more to specify the size of the reference device, in millimeters
        var deviceSizeInMlmX = reader.ReadUInt32();
        var deviceSizeInMlmY = reader.ReadUInt32();

        var widthInPixel = Math.Round(0.5 + (right - left + 1.0) * deviceSizeInPixelX / deviceSizeInMlmX / 100.0);
        var heightInPixel = Math.Round(0.5 + (bottom-top + 1.0) * deviceSizeInPixelY / deviceSizeInMlmY / 100.0);

        if (widthInPixel <= 0 || heightInPixel <= 0)
            return Size.Empty;

        return new Size((uint) widthInPixel, (uint) heightInPixel);
    }

    private static Size DecodeXml(Stream stream)
    {
        try
        {
            var nav = new XPathDocument(stream).CreateNavigator();
            // use local-name() to ignore any xml namespace
            nav = nav.SelectSingleNode("/*[local-name() = 'svg']");
            if (nav is not null)
            {
                var width = Unit.Parse(nav.GetAttribute("width", string.Empty).AsSpan());
                var height = Unit.Parse(nav.GetAttribute("height", string.Empty).AsSpan());
                if (width.IsValid && height.IsValid)
                    return new Size((uint) width.ValueInPx, (uint) height.ValueInPx);
            }
        }
        catch (SystemException)
        {
            return Size.Empty;
        }

        return Size.Empty;
    }
}
