/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
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