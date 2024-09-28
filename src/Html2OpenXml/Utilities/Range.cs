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
namespace System;

#if !NET5_0_OR_GREATER
readonly struct Range(int start, int end)
{
    /// <summary>Represent the inclusive start index of the Range.</summary>
    public int Start { get; } = start;

    /// <summary>Represent the exclusive end index of the Range.</summary>
    public int End { get; } = end;
}
#endif