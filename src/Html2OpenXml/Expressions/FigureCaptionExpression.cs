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
using System.Globalization;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a <c>figcaption</c> element, which is used to describe an image.
/// </summary>
sealed class FigureCaptionExpression(IHtmlElement node) : PhrasingElementExpression(node)
{

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        ComposeStyles(context);
        var childElements = Interpret(context.CreateChild(this), node.ChildNodes);
        if (!childElements.Any())
            return [];

        var p = new Paragraph (
            new Run(
                new Text("Figure ") { Space = SpaceProcessingModeValues.Preserve }
            ),
            new SimpleField(
                new Run(
                    new Text(AddFigureCaption(context).ToString(CultureInfo.InvariantCulture)))
            ) { Instruction = " SEQ Figure \\* ARABIC " }
        ) {
            ParagraphProperties = new ParagraphProperties {
                ParagraphStyleId = context.DocumentStyle.GetParagraphStyle(context.DocumentStyle.DefaultStyles.CaptionStyle),
                KeepNext = new KeepNext()
            }
        };

        if (childElements.First() is Run run) // any caption?
        {
            Text? t = run.GetFirstChild<Text>();
            if (t != null)
                t.Text = " " + t.InnerText; // append a space after the numero of the picture
        }

        return [p];
    }

    /// <summary>
    /// Add a new figure caption to the document.
    /// </summary>
    /// <returns>Returns the id of the new figure caption.</returns>
    private static int AddFigureCaption(ParsingContext context)
    {
        var figCaptionRef = context.Properties<int?>("figCaptionRef");
        if (!figCaptionRef.HasValue)
        {
            figCaptionRef = 0;
            foreach (var p in context.MainPart.Document.Descendants<SimpleField>())
            {
                if (p.Instruction == " SEQ Figure \\* ARABIC ")
                    figCaptionRef++;
            }
        }
        figCaptionRef++;

        context.Properties("figCaptionRef", figCaptionRef);
        return figCaptionRef.Value;
    }
}