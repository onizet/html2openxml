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
using System.Collections.Generic;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Leaf expression which process a simple text content.
/// </summary>
sealed class TextExpression(INode node) : HtmlDomExpression
{
    private readonly INode node = node;

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        string text = node.TextContent.Normalize();
        if (text.Trim().Length == 0) return [];

        if (!context.PreserveLinebreaks)
            text = text.CollapseLineBreaks();
        if (context.CollapseWhitespaces && text[0].IsWhiteSpaceCharacter() &&
            node.PreviousSibling is IHtmlImageElement)
        {
            text = " " + text.CollapseAndStrip();
        }
        else if (context.CollapseWhitespaces)
            text = text.CollapseAndStrip();

        if (!context.PreserveLinebreaks)
            return [new Run(new Text(text))];

        var run = new Run();
        char[] chars = text.ToCharArray();
        int shift = 0, c = 0;
        bool wasCR = false; // avoid adding 2 breaks for \r\n
        for ( ; c < chars.Length ; c++)
        {
            if (!chars[c].IsLineBreak())
            {
                wasCR = false;
                continue;
            }

            if (wasCR) continue;
            wasCR = chars[c] == Symbols.CarriageReturn;

            if (c > 1)
            {
                run.Append(new Text(new string(chars, shift, c - shift)) 
                    { Space = SpaceProcessingModeValues.Preserve });
                run.Append(new Break());
            }
            shift = c + 1;
        }

        if (c > shift)
            run.Append(new Text(new string(chars, shift, c - shift)) 
                { Space = SpaceProcessingModeValues.Preserve });

        return [run];
    }
}
