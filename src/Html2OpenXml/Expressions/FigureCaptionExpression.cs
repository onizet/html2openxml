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
using System.Globalization;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a <c>figcaption</c> element, which is used to describe an image.
/// </summary>
sealed class FigureCaptionExpression(IHtmlElement node) : BlockElementExpression(node)
{

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        ComposeStyles(context);
        var childElements = Interpret(context.CreateChild(this), node.ChildNodes);

        var figNumRef = new List<OpenXmlElement>() {
            new Run(
                new Text("Figure ") { Space = SpaceProcessingModeValues.Preserve }
            ),
            new SimpleField(
                new Run(
                    new Text(AddFigureCaption(context).ToString(CultureInfo.InvariantCulture)))
            ) { Instruction = " SEQ Figure \\* ARABIC " }
        };


        if (!childElements.Any())
        {
            return [new Paragraph(figNumRef) {
                ParagraphProperties = new ParagraphProperties {
                    ParagraphStyleId = context.DocumentStyle.GetParagraphStyle(context.DocumentStyle.DefaultStyles.CaptionStyle),
                    KeepNext = DetermineKeepNext(node),
                }
            }];
        }

        //Add the figure number references to the start of the first paragraph.
        if(childElements.First() is Paragraph p)
        {
           var properties = p.GetFirstChild<ParagraphProperties>();
           p.InsertAfter(new Run(
              new Text(" ") { Space = SpaceProcessingModeValues.Preserve }
           ), properties);
           p.InsertAfter(figNumRef[1], properties);
           p.InsertAfter(figNumRef[0], properties);
        }
        else
        {
            // The first child of the figure caption is a table or something.
            // Just prepend a new paragraph with the figure number reference.
            childElements =  [
                new Paragraph(figNumRef),
                ..childElements
            ];
        }

        foreach (var paragraph in childElements.OfType<Paragraph>())
        {
            paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph.ParagraphProperties.ParagraphStyleId ??= context.DocumentStyle.GetParagraphStyle(context.DocumentStyle.DefaultStyles.CaptionStyle);
            //Keep caption paragraphs together.
            paragraph.ParagraphProperties.KeepNext = new KeepNext();
        }

        if(childElements.OfType<Paragraph>().LastOrDefault() is Paragraph lastPara)
        {
            lastPara.ParagraphProperties!.KeepNext = DetermineKeepNext(node);
        }

        return childElements;
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
            foreach (var p in context.MainPart.Document!.Descendants<SimpleField>())
            {
                if (p.Instruction == " SEQ Figure \\* ARABIC ")
                    figCaptionRef++;
            }
        }
        figCaptionRef++;

        context.Properties("figCaptionRef", figCaptionRef);
        return figCaptionRef.Value;
    }

    /// <summary>
    /// Determines whether the KeepNext property should apply this this caption.
    /// </summary>
    /// <returns>A new <see cref="KeepNext"/> or null.</returns>
    private static KeepNext? DetermineKeepNext(IHtmlElement node)
    {
        // A caption at the end of a figure will have no next sibling.
        if(node.NextElementSibling is null)
        {
            return null;
        }
        return new();
    }
}
