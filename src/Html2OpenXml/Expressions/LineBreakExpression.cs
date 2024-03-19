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
/// Process the parsing of a <c>br</c> element.
/// </summary>
sealed class LineBreakExpression(IHtmlElement node) : HtmlElementExpression(node)
{
    public override void CascadeStyles(OpenXmlCompositeElement element)
    {
        throw new System.NotSupportedException();
    }

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlCompositeElement> Interpret (ParsingContext context)
    {
        return [new Run(new Break())];
    }
}