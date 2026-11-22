/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml
 * Licensed under the MIT License.
 * Please see the LICENSE file in the project root fo full license information
 */
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
        var childElements = Interpret(context.CreateChild(this), node.ChildNodes).ToList();
        if (childElements.Count == 0)
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

        if (childElements[0] is Run run) // any caption?
        {
            Text? t = run.GetFirstChild<Text>();
            t?.Text = " " + t.InnerText;
        }
        p.Append(childElements);

        var att = styleAttributes["text-align"];
        if (att.IsEmpty) att = node.GetAttribute("align").AsSpan();
        if (!att.IsEmpty)
        {
            if (Converter.TryParseHAlign(att, out var align))
                p.ParagraphProperties.Justification = new() { Val = align };
        }
        else
        {
            // If no particular alignment has been specified for the legend, we will align the legend
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