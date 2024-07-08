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
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a <c>caption</c> element, which is used to describe a table.
/// </summary>
sealed class TableCaptionExpression(Table table, IHtmlElement node) : PhrasingElementExpression(node)
{
    private readonly Table table = table;

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        ComposeStyles(context);
        var childElements = Interpret(context.CreateChild(this), node.ChildNodes);
        if (!childElements.Any())
            return [];

        var p = new Paragraph (
            new Run(
                new FieldChar() { FieldCharType = FieldCharValues.Begin }),
            new Run(
                new FieldCode("SEQ TABLE \\* ARABIC") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(
                new FieldChar() { FieldCharType = FieldCharValues.End })
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
                t.Text = " " + t.InnerText;
        }
        p.Append(childElements);

        string? att = styleAttributes!["text-align"] ?? node.GetAttribute("align");
        if (!string.IsNullOrEmpty(att))
        {
            JustificationValues? align = Converter.ToParagraphAlign(att);
            if (align.HasValue)
                p.ParagraphProperties.Justification = new() { Val = align };
        }
        else
        {
            // If no particular alignement has been specified for the legend, we will align the legend
            // relative to the owning table
            TableProperties? props = table.GetFirstChild<TableProperties>();
            if (props != null)
            {
                TableJustification? justif = props.GetFirstChild<TableJustification>();
                if (justif?.Val != null) 
                    p.ParagraphProperties.Justification = new() { Val = justif.Val.Value.ToJustification() };
            }
        }

        return [p];
    }
}