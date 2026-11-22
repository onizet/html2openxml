/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
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