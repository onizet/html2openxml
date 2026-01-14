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
#if !NET5_0_OR_GREATER
namespace System;

using System.Runtime.CompilerServices;

readonly struct Range(int start, int end)
{
    /// <summary>Represent the inclusive start index of the Range.</summary>
    public int Start { get; } = start;

    /// <summary>Represent the exclusive end index of the Range.</summary>
    public int End { get; } = end;

    /// <summary>Calculate the start offset and length of range object using a collection length.</summary>
    /// <remarks>
    /// For performance reason, we don't validate the input length parameter against negative values.
    /// It is expected Range will be used with collections which always have non negative length/count.
    /// We validate the range is inside the length scope though.
    /// </remarks>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public (int Offset, int Length) GetOffsetAndLength(int _)
    {
        return (Start, End - Start);
    }
}
#endif