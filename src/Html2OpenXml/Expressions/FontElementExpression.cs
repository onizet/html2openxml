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
using System;
using System.Globalization;
using AngleSharp.Html.Dom;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a <c>font</c> element.
/// </summary>
sealed class FontElementExpression(IHtmlElement node) : PhrasingElementExpression(node)
{
    protected override void ComposeStyles(ParsingContext context)
    {
        base.ComposeStyles(context);

        string? attrValue = node.GetAttribute("size");
        if (!string.IsNullOrEmpty(attrValue))
        {
            Unit fontSize = Converter.ToFontSize(attrValue);
            if (fontSize.IsFixed)
                runProperties.FontSize = new() { 
                    Val = Math.Round(fontSize.ValueInPoint * 2).ToString(CultureInfo.InvariantCulture) };
        }

        attrValue = node.GetAttribute("face");
        if (!string.IsNullOrEmpty(attrValue))
        {
            // Set HightAnsi. Bug fixed by xjpmauricio on github.com/onizet/html2openxml/discussions/285439
            // where characters with accents where always using fallback font
            runProperties.RunFonts = new() { Ascii = attrValue, HighAnsi = attrValue };
        }
    }
}