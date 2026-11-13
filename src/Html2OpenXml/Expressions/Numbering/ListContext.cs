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
using AngleSharp.Dom;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Represents the current numbering state while traversing nested HTML lists.
/// </summary>
readonly record struct ListContext(string Name, int AbsNumId, int InstanceId,
    int Level, DirectionMode? Dir, long MarginLeftDxa)
{
    public ListContext NextLevel(DirectionMode? dir, long extraMargin)
        => this with {
            Level = Level+1,
            Dir = dir ?? Dir,
            MarginLeftDxa = MarginLeftDxa + extraMargin
        };
}