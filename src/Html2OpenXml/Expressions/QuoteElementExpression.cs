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
/// Process the parsing of <c>quote</c> element.
/// </summary>
sealed class QuoteElementExpression(IHtmlElement node) : PhrasingElementExpression(node)
{

    public override IEnumerable<OpenXmlElement> Interpret(ParsingContext context)
    {
        // The browsers render the quote tag between a kind of separators.
        // We add the Quote style to the nested runs to match more Word.

        Run prefixRun = new(
            new Text(" " + context.DocumentStyle.QuoteCharacters.Prefix) { Space = SpaceProcessingModeValues.Preserve }
        );
        prefixRun.RunProperties = runProperties;
        prefixRun.RunProperties.RunStyle = context.DocumentStyle.GetRunStyle(context.DocumentStyle.DefaultStyles.QuoteStyle);

        yield return prefixRun;
        var elements = base.Interpret(context);
        foreach (var el in elements)
            yield return el;

        Run suffixRun = new(
            new Text(context.DocumentStyle.QuoteCharacters.Suffix) { Space = SpaceProcessingModeValues.Preserve }
        );
        suffixRun.RunProperties = (RunProperties) runProperties.CloneNode(true);
        suffixRun.RunProperties.RunStyle = context.DocumentStyle.GetRunStyle(context.DocumentStyle.DefaultStyles.QuoteStyle);

        yield return suffixRun;
    }
}