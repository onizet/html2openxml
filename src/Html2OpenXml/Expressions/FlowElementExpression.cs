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
using System;
using System.Collections.Generic;
using System.Globalization;
using AngleSharp.Html.Dom;
using AngleSharp.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of flow contents. Flow content are sectioning tags, body, heading and footer tags.
/// </summary>
class FlowElementExpression(IHtmlElement node) : PhrasingElementExpression(node)
{
    protected readonly ParagraphProperties paraProperties = new();


    /// <inheritdoc/>
    public override IEnumerable<OpenXmlCompositeElement> Interpret (ParsingContext context)
    {
        //TODO: add break? elements.Add(new Run(new Break()));
        return base.Interpret(context);
    }

    protected override IEnumerable<OpenXmlCompositeElement> Interpret (
        ParsingContext context, IEnumerable<AngleSharp.Dom.INode> childNodes)
    {
        var runs = new List<Run>();
        var flowElements = new List<OpenXmlCompositeElement>();

        if ("always".Equals(styleAttributes!["page-break-before"], StringComparison.OrdinalIgnoreCase))
        {
            runs.Add(
                new Run(
                    new Break() { Type = BreakValues.Page })
            );
            runs.Add(new Run(
                new LastRenderedPageBreak())
            );
        }

        foreach (var child in childNodes)
        {
            var expression = CreateFromHtmlNode (child);
            if (expression == null) continue;

            foreach (var element in expression.Interpret(context))
            {
                context.CascadeStyles(element);
                if (element is Run r)
                {
                    runs.Add(r);
                    continue;
                }

                if (runs.Count > 0)
                {
                    flowElements.Add(CombineRuns(runs));
                    runs.Clear();
                }

                flowElements.Add(element);
            }
        }

        if ("always".Equals(styleAttributes!["page-break-after"], StringComparison.OrdinalIgnoreCase))
        {
            runs.Add(new Run(
                new Break() { Type = BreakValues.Page }));
        }

        if (runs.Count > 0)
            flowElements.Add(CombineRuns(runs));

        return flowElements;
    }

    public override void CascadeStyles(OpenXmlCompositeElement element)
    {
        base.CascadeStyles(element);
        if (!paraProperties.HasChildren || element is not Paragraph paragraph)
            return;

        paragraph.ParagraphProperties ??= new ParagraphProperties();

        var knownTags = new HashSet<string>();
        foreach (var prop in paragraph.ParagraphProperties)
        {
            if (!knownTags.Contains(prop.LocalName))
                knownTags.Add(prop.LocalName);
        }

        foreach (var prop in paraProperties)
        {
            if (!knownTags.Contains(prop.LocalName))
                paragraph.ParagraphProperties.AddChild(prop.CloneNode(true));
        }
    }

    /// <inheritdoc/>
    protected override void ComposeStyles (ParsingContext context)
    {
        base.ComposeStyles(context);

        if (node.Language != null && node.Language != node.Owner!.Body!.Language)
        {
            try
            {
                var ci = new CultureInfo(node.Language);
                bool rtl = ci.TextInfo.IsRightToLeft;

                var lang = new Languages() { Val = ci.TwoLetterISOLanguageName };
                if (rtl) lang.Bidi = ci.Name;

                paraProperties.ParagraphMarkRunProperties = new ParagraphMarkRunProperties(lang);
                paraProperties.BiDi = new BiDi() { Val = OnOffValue.FromBoolean(rtl) };
            }
            catch (ArgumentException)
            {
                // lang not valid, ignore it
            }
        }


        var attrValue = styleAttributes!["text-align"];
        JustificationValues? align = Converter.ToParagraphAlign(attrValue);
        if (align.HasValue)
        {
            paraProperties.Justification = new() { Val = align };
        }

        // according to w3c, dir should be used in conjonction with lang. But whatever happens, we'll apply the RTL layout
        if ("rtl".Equals(node.Direction, StringComparison.OrdinalIgnoreCase))
        {
            paraProperties.Justification = new() { Val = JustificationValues.Right };
        }
        else if ("ltr".Equals(node.Direction, StringComparison.OrdinalIgnoreCase))
        {
            paraProperties.Justification = new() { Val = JustificationValues.Left };
        }

        var border = styleAttributes.GetAsBorder();
        if (!border.IsEmpty)
        {
            ParagraphBorders borders = new();
            if (border.Top.IsValid) borders.TopBorder = 
                new() { Val = border.Top.Style, Color = border.Top.Color.ToHexString(), Size = (uint) border.Top.Width.ValueInPoint, Space = 1U };
            if (border.Left.IsValid) borders.LeftBorder =
                new() { Val = border.Left.Style, Color = border.Left.Color.ToHexString(), Size = (uint) border.Left.Width.ValueInPoint, Space = 1U };
            if (border.Bottom.IsValid) borders.BottomBorder =
                new() { Val = border.Bottom.Style, Color = border.Bottom.Color.ToHexString(), Size = (uint) border.Bottom.Width.ValueInPoint, Space = 1U };
            if (border.Right.IsValid) borders.RightBorder =
                new() { Val = border.Right.Style, Color = border.Right.Color.ToHexString(), Size = (uint) border.Right.Width.ValueInPoint, Space = 1U };

            paraProperties.ParagraphBorders = borders;
        }

        foreach (string className in node.ClassList)
        {
            var matchClassName = context.DocumentStyle.GetStyle(className, StyleValues.Paragraph, ignoreCase: true);
            if (matchClassName != null)
            {
                paraProperties.ParagraphStyleId = new ParagraphStyleId() { Val = matchClassName };
                break;
            }
        }

        Margin margin = styleAttributes.GetAsMargin("margin");
        Indentation? indentation = null;
        if (!margin.IsEmpty)
        {
            if (margin.Top.IsFixed || margin.Bottom.IsFixed)
            {
                var spacing = new SpacingBetweenLines();
                if (margin.Top.IsFixed) spacing.Before = margin.Top.ValueInDxa.ToString(CultureInfo.InvariantCulture);
                if (margin.Bottom.IsFixed) spacing.After = margin.Bottom.ValueInDxa.ToString(CultureInfo.InvariantCulture);
                paraProperties.SpacingBetweenLines = spacing;
            }
            if (margin.Left.IsFixed || margin.Right.IsFixed)
            {
                indentation = new Indentation();
                if (margin.Left.IsFixed) indentation.Left = margin.Left.ValueInDxa.ToString(CultureInfo.InvariantCulture);
                if (margin.Right.IsFixed) indentation.Right = margin.Right.ValueInDxa.ToString(CultureInfo.InvariantCulture);
                paraProperties.Indentation = indentation;
            }
        }

        // implemented by giorand (feature #13787)
        Unit textIndent = styleAttributes.GetAsUnit("text-indent");
        if (textIndent.IsValid)
        {
            indentation ??= new Indentation();
            indentation.FirstLine = textIndent.ValueInDxa.ToString(CultureInfo.InvariantCulture);
            paraProperties.Indentation = indentation;
        }
    }

    /// <summary>
    /// Mimics the behaviour of Html rendering when 2 consecutives runs are separated by a space
    /// </summary>
    private static Paragraph CombineRuns(IList<Run> runs)
    {
        if (runs.Count == 1)
            return new Paragraph(runs);

        var p = new Paragraph();
        bool endsWithSpace = true;
        foreach (var run in runs)
        {
            var textElement = run.GetFirstChild<Text>()!;
            if (textElement != null) // could be null when <br/>
            {
                var text = textElement.Text;
                // we know that the text cannot be empty because in TextExpression,
                // we skip them
                if (!endsWithSpace && !text[0].IsSpaceCharacter())
                {
                    textElement.Text = " " + text;
                }
                endsWithSpace = text[text.Length - 1].IsSpaceCharacter();
            }
            p.AppendChild(run);
        }

        return p;
    }
}