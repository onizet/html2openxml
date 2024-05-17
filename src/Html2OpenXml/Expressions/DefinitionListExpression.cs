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
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the definition list item <c>dl</c>.
/// </summary>
class DefinitionListExpression(IHtmlElement node) : BlockElementExpression(node)
{
    public override IEnumerable<OpenXmlCompositeElement> Interpret(ParsingContext context)
    {
        var childElements = base.Interpret(context);

        var paragraph = new Paragraph (childElements);
        paragraph.ParagraphProperties = new ParagraphProperties() {
            Indentation = new() { FirstLine = "708" },
            SpacingBetweenLines = new() { After = "0" }
        };
        CascadeStyles(paragraph);
        return [paragraph];
    }
}