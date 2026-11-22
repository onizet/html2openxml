/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of <c>quote</c> element.
/// </summary>
sealed class QuoteElementExpression(IHtmlElement node) : PhrasingElementExpression(node)
{

    public override IEnumerable<OpenXmlElement> Interpret(ParsingContext context)
    {
        // The browsers render the quote tag between a kind of separators.
        // We apply the Quote character style to better match Word rendering.

        Run prefixRun = new(
            new Text(" " + context.DocumentStyle.QuoteCharacters.Prefix) { Space = SpaceProcessingModeValues.Preserve }
        );
        prefixRun.RunProperties = runProperties;
        prefixRun.RunProperties.RunStyle = context.DocumentStyle.GetRunStyle(context.DocumentStyle.DefaultStyles.QuoteStyle);

        yield return prefixRun;
        foreach (var el in base.Interpret(context))
            yield return el;

        Run suffixRun = new(
            new Text(context.DocumentStyle.QuoteCharacters.Suffix) { Space = SpaceProcessingModeValues.Preserve }
        );
        suffixRun.RunProperties = (RunProperties) runProperties.CloneNode(true);
        suffixRun.RunProperties.RunStyle = context.DocumentStyle.GetRunStyle(context.DocumentStyle.DefaultStyles.QuoteStyle);

        yield return suffixRun;
    }
}