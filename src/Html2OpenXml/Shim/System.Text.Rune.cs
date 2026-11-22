/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */

#if !NET5_0_OR_GREATER

using System.Runtime.CompilerServices;

namespace System.Text;

readonly record struct Rune(uint Value)
{
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool IsValid(uint codePoint)
    {
        return codePoint <= 0x10FFFF &&
            (codePoint < 0xD800 || codePoint > 0xDFFF);
    }

    /// <summary>
    /// Encodes this Rune to a UTF-16 destination buffer.
    /// </summary>
    /// <returns>The number of <see cref="Char"/> values written to <c>destination</c>.</returns>
    public int EncodeToUtf16(Span<char> destination)
    {
        if (Value < 0xFFFFu)
        {
            if (destination.Length < 1)
                throw new ArgumentException("Destination is too short.");

            destination[0] = (char) Value;
            return 1;
        }

        if (destination.Length < 2)
            throw new ArgumentException("Destination is too short.");

        destination[0] = (char) ((Value + ((0xD800u - 0x40u) << 10)) >> 10);
        destination[1] = (char) ((Value & 0x3FFu) + 0xDC00u);
        return 2;
    }
}
#endif
