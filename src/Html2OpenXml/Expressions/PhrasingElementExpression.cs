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
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a phrasing content. A Phrasing content is an inline layout content at the lower level
/// that consists of text and HTML elements that mark up the text within paragraphs.
/// </summary>
class PhrasingElementExpression(IHtmlElement node, OpenXmlLeafElement? styleProperty = null) : HtmlElementExpression
{
    private readonly OpenXmlLeafElement? defaultStyleProperty = styleProperty;

    protected readonly RunProperties runProperties = new();
    protected HtmlAttributeCollection? styleAttributes;
    protected IHtmlElement node = node;


    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        ComposeStyles(context);
        return Interpret(context.CreateChild(this), node.ChildNodes);
    }

    protected virtual IEnumerable<OpenXmlElement> Interpret (
        ParsingContext context, IEnumerable<INode> childNodes)
    {
        var runs = new List<OpenXmlElement>();
        foreach (var child in childNodes)
        {
            var expression = CreateFromHtmlNode (child);
            if (expression == null) continue;

            foreach (var element in expression.Interpret(context))
            {
                context.CascadeStyles(element);
                runs.Add(element);
            }
        }
        return CombineRuns(runs);
    }

    public override void CascadeStyles(OpenXmlElement element)
    {
        if (!runProperties.HasChildren || element is not Run run)
            return;

        run.RunProperties ??= new();

        var knownTags = new HashSet<string>();
        foreach (var prop in run.RunProperties)
        {
            if (!knownTags.Contains(prop.LocalName))
                knownTags.Add(prop.LocalName);
        }

        foreach (var prop in runProperties)
        {
            if (!knownTags.Contains(prop.LocalName))
                run.RunProperties.AddChild(prop.CloneNode(true));
        }
    }

    /// <summary>
    /// Prepare the conversion of style attributes of the current Html Dom element to OpenXml equivalent.
    /// </summary>
    protected virtual void ComposeStyles (ParsingContext context)
    {
        styleAttributes = node.GetStyles();
        if (defaultStyleProperty != null)
            runProperties.AddChild(defaultStyleProperty.CloneNode(true));

        if (node.Language != null && node.Language != node.Owner!.Body!.Language)
        {
            var ci = Converter.ToLanguage(node.Language);
            if (ci != null)
            {
                runProperties.Languages = new () { Val = ci.TwoLetterISOLanguageName };
            }
        }

        // according to w3c, dir should be used in conjonction with lang. But whatever happens, we'll apply the RTL layout
        var dir = node.GetTextDirection();
        if (dir == DirectionMode.Rtl)
        {
            runProperties.RightToLeftText = new RightToLeftText();
        }

        // OpenXml limits the border to 4-side of the same color and style.
        SideBorder border = styleAttributes.GetSideBorder("border");
        if (border.IsValid)
        {
            runProperties.Border = new Border() {
                Val = border.Style,
                Color = border.Color.ToHexString(),
                Size = (uint) border.Width.ValueInPx * 4,
                Space = 1U
            };
        }

        var colorValue = styleAttributes.GetColor("color");
        if (colorValue.IsEmpty) colorValue = HtmlColor.Parse(node.GetAttribute("color"));
        if (!colorValue.IsEmpty)
            runProperties.Color = new Color { Val = colorValue.ToHexString() };

        var bgcolor = styleAttributes.GetColor("background-color");
        if (bgcolor.IsEmpty) bgcolor = styleAttributes.GetColor("background");
        if (!bgcolor.IsEmpty)
        {
            // change the way the background-color renders. It now uses Shading instead of Highlight.
            // Changes brought by Wude on http://html2openxml.codeplex.com/discussions/277570
            runProperties.Shading = new Shading { Val = ShadingPatternValues.Clear, Fill = bgcolor.ToHexString() };
        }

        foreach (var decoration in Converter.ToTextDecoration(styleAttributes["text-decoration"]))
        {
            switch (decoration)
            {
                case TextDecoration.Underline:
                    runProperties.Underline = new Underline { Val = UnderlineValues.Single }; break;
                case TextDecoration.Dotted:
                    runProperties.Underline = new Underline { Val = UnderlineValues.Dotted }; break;
                case TextDecoration.Dashed:
                    runProperties.Underline = new Underline { Val = UnderlineValues.Dash }; break;
                case TextDecoration.Wave:
                    runProperties.Underline = new Underline { Val = UnderlineValues.Wave }; break;
                case TextDecoration.Double:
                    runProperties.DoubleStrike = new DoubleStrike(); break;
                case TextDecoration.LineThrough:
                    runProperties.Strike = new Strike(); break;
            }
        }

        // these style cannot be defined at the same time
        if (runProperties.DoubleStrike != null)
            runProperties.Strike = null;

        foreach(string className in node.ClassList)
        {
            var matchClassName = context.DocumentStyle.GetStyle(className, StyleValues.Character, ignoreCase: true);
            // only one Style can be applied in OpenXml and dealing with inheritance is out of scope
            if (matchClassName != null)
            {
                runProperties.RunStyle = new RunStyle() { Val = matchClassName };
                break;
            }
        }

        HtmlFont font = styleAttributes.GetFont("font");

        if (font.Style == FontStyle.Italic)
            runProperties.Italic = new Italic();
        else if (font.Style == FontStyle.Normal)
            runProperties.Italic = new Italic() { Val = false };

        if (font.Weight == FontWeight.Bold || font.Weight == FontWeight.Bolder)
            runProperties.Bold = new Bold();
        else if (font.Weight == FontWeight.Normal)
            runProperties.Bold = new Bold() { Val = false };

        if (font.Variant == FontVariant.SmallCaps)
            runProperties.SmallCaps = new SmallCaps();
        else if (font.Variant == FontVariant.Normal)
            runProperties.SmallCaps = new SmallCaps() { Val = false };

        if (font.Family != null)
            runProperties.RunFonts = new RunFonts() { Ascii = font.Family, HighAnsi = font.Family };

        // size are half-point font size
        if (font.Size.IsFixed)
            runProperties.FontSize = new FontSize() { Val = Math.Round(font.Size.ValueInPoint * 2).ToString(CultureInfo.InvariantCulture) };
    }

    /// <summary>
    /// Mimics the behaviour of Html rendering when 2 consecutives runs are separated by a space.
    /// </summary>
    protected static IEnumerable<OpenXmlElement> CombineRuns(IEnumerable<OpenXmlElement> runs)
    {
        if (runs.Count() == 1)
        {
            yield return runs.First();
            yield break;
        }

        bool endsWithSpace = true;
        foreach (var run in runs)
        {
            var textElement = run.GetFirstChild<Text>();
            // run can be also a hyperlink
            textElement ??= run.GetFirstChild<Run>()?.GetFirstChild<Text>();

            if (textElement != null) // could be null when <br/>
            {
                var text = textElement.Text;
                // we know that the text cannot be empty because we skip them in TextExpression
                if (!endsWithSpace && !text[0].IsSpaceCharacter())
                {
                    yield return new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve });
                }
                endsWithSpace = text[text.Length - 1].IsSpaceCharacter();
            }
            else if (run.LastChild is Break)
            {
                endsWithSpace = true;
            }
            yield return run;
        }
    }
}