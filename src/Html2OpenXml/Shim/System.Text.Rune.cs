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
 */
using System.Runtime.CompilerServices;

#if !NET5_0_OR_GREATER

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
