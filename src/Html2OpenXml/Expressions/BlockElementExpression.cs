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
using System.Linq;
using AngleSharp.Html.Dom;
using AngleSharp.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of block contents (like <c>p</c>, <c>span</c>, <c>heading</c>).
/// A block-level element always starts on a new line, and the browsers automatically add some space (a margin) before and after the element.
/// </summary>
class BlockElementExpression(IHtmlElement node) : PhrasingElementExpression(node)
{
    protected readonly ParagraphProperties paraProperties = new();


    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        //TODO: add break? elements.Add(new Run(new Break()));
        var elements = base.Interpret(context);

        var isBookmarkTarget = node.GetAttribute(InternalNamespaceUri, "bookmark");
        if (isBookmarkTarget is not null)
        {
            elements.First().PrependChild(new BookmarkStart() { Name = node.Id });
            elements.First().AppendChild(new BookmarkEnd());
        }

        return elements;
    }

    protected override IEnumerable<OpenXmlElement> Interpret (
        ParsingContext context, IEnumerable<AngleSharp.Dom.INode> childNodes)
    {
        var runs = new List<Run>();
        var flowElements = new List<OpenXmlElement>();

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
                    flowElements.Add(CombineRuns(context, runs, paraProperties));
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
            flowElements.Add(CombineRuns(context, runs, paraProperties));

        return flowElements;
    }

    public override void CascadeStyles(OpenXmlElement element)
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


        var styleBorder = styleAttributes.GetBorders();
        if (!styleBorder.IsEmpty)
        {
            var borders = new ParagraphBorders {
                LeftBorder = Converter.ToBorder<LeftBorder>(styleBorder.Left),
                RightBorder = Converter.ToBorder<RightBorder>(styleBorder.Right),
                TopBorder = Converter.ToBorder<TopBorder>(styleBorder.Top),
                BottomBorder = Converter.ToBorder<BottomBorder>(styleBorder.Bottom)
            };

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

        Margin margin = styleAttributes.GetMargin("margin");
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
        Unit textIndent = styleAttributes.GetUnit("text-indent");
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
    internal static Paragraph CombineRuns(ParsingContext context, IList<Run> runs, ParagraphProperties paraProperties)
    {
        Paragraph p = new();
        if (paraProperties.HasChildren)
            p.ParagraphProperties = (ParagraphProperties) paraProperties.CloneNode(true);

        context.CascadeStyles(p);

        if (runs.Count == 1)
        {
            p.AddChild(runs.First());
            return p;
        }

        bool endsWithSpace = true;
        foreach (var run in runs)
        {
            var textElement = run.GetFirstChild<Text>()!;
            if (textElement != null) // could be null when <br/>
            {
                var text = textElement.Text;
                // we know that the text cannot be empty because we skip them in TextExpression
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