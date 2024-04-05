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
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a phrasing content. A Phrasing content is the content at the lower level
/// that consists of text and HTML elements that mark up the text within paragraphs.
/// </summary>
class PhrasingElementExpression(IHtmlElement node, OpenXmlLeafElement? styleProperty = null) : HtmlElementExpression(node)
{
    private readonly OpenXmlLeafElement? defaultStyleProperty = styleProperty;

    protected readonly RunProperties runProperties = new();


    /// <inheritdoc/>
    public override IEnumerable<OpenXmlCompositeElement> Interpret (ParsingContext context)
    {
        ComposeStyles(context);
        return Interpret(context.CreateChild(this), node.ChildNodes);
    }

    protected virtual IEnumerable<OpenXmlCompositeElement> Interpret (
        ParsingContext context, IEnumerable<INode> childNodes)
    {
        foreach (var child in childNodes)
        {
            var expression = CreateFromHtmlNode (child);
            if (expression == null) continue;

            foreach (var element in expression.Interpret(context))
            {
                context.CascadeStyles(element);
                yield return element;
            }
        }
    }

    public override void CascadeStyles(OpenXmlCompositeElement element)
    {
        if (!runProperties.HasChildren || element is not Run run)
            return;

        run.RunProperties ??= new RunProperties();

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
        var styleAttributes = HtmlAttributeCollection.ParseStyle(node.GetAttribute("style"));
        if (defaultStyleProperty != null)
            runProperties.AddChild(defaultStyleProperty.CloneNode(true));

        if (!string.IsNullOrWhiteSpace(node.Language))
        {
            try
            {
                var ci = new CultureInfo(node.Language);
                Languages lang = new () { Val = ci.TwoLetterISOLanguageName };
                runProperties.Languages = new () { Bidi = ci.Name };
            }
            catch (ArgumentException)
            {
                // lang not valid, ignore it
            }
        }

        // according to w3c, dir should be used in conjonction with lang. But whatever happens, we'll apply the RTL layout
        if ("rtl".Equals(node.Direction, StringComparison.OrdinalIgnoreCase))
        {
            runProperties.RightToLeftText = new RightToLeftText();
        }

        // OpenXml limits the border to 4-side of the same color and style.
        SideBorder border = styleAttributes.GetAsSideBorder("border");
        if (border.IsValid)
        {
            runProperties.Border = new Border() {
                Val = border.Style,
                Color = border.Color.ToHexString(),
                Size = (uint) border.Width.ValueInPx * 4,
                Space = 1U
            };
        }

        var colorValue = styleAttributes.GetAsColor("color");
        if (colorValue.IsEmpty) colorValue = HtmlColor.Parse(node.GetAttribute("color"));
        if (!colorValue.IsEmpty)
            runProperties.Color = new Color { Val = colorValue.ToHexString() };

        colorValue = styleAttributes.GetAsColor("background-color");
        if (!colorValue.IsEmpty)
        {
            // change the way the background-color renders. It now uses Shading instead of Highlight.
            // Changes brought by Wude on http://html2openxml.codeplex.com/discussions/277570
            runProperties.Shading = new Shading { Val = ShadingPatternValues.Clear, Fill = colorValue.ToHexString() };
        }

        var decorations = Converter.ToTextDecoration(styleAttributes["text-decoration"]);
        if ((decorations & TextDecoration.Underline) != 0)
        {
            runProperties.Underline = new Underline { Val = UnderlineValues.Single };
        }
        if ((decorations & TextDecoration.LineThrough) != 0)
        {
            runProperties.Strike = new Strike();
        }

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

        HtmlFont font = styleAttributes.GetAsFont("font");

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
}