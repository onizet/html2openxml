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
using AngleSharp.Html.Dom;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of <c>cite</c> element.
/// </summary>
sealed class CiteElementExpression(IHtmlElement node) : PhrasingElementExpression(node)
{
    protected override void ComposeStyles(ParsingContext context)
    {
        base.ComposeStyles(context);
        runProperties.RunStyle ??= context.DocumentStyle.GetRunStyle(context.DocumentStyle.DefaultStyles.QuoteStyle);
    }
}