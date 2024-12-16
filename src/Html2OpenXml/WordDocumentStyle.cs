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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml;

/// <summary>
/// Defines the styles to apply on OpenXml elements.
/// </summary>
public sealed class WordDocumentStyle
{
    /// <summary>
    /// Occurs when a Style is missing in the MainDocumentPart but will be used during the conversion process.
    /// </summary>
    public event EventHandler<StyleEventArgs>? StyleMissing;

    private readonly MainDocumentPart mainPart;
    private readonly OpenXmlDocumentStyleCollection knownStyles = [];
    private readonly HashSet<string> lazyPredefinedStyles;

    private DefaultStyles? defaultStyles;
    

    internal WordDocumentStyle(MainDocumentPart mainPart)
    {
        PrepareStyles(mainPart);
        lazyPredefinedStyles = [ 
            PredefinedStyles.Caption,
            PredefinedStyles.EndnoteReference,
            PredefinedStyles.EndnoteText,
            PredefinedStyles.FootnoteReference,
            PredefinedStyles.FootnoteText,
            PredefinedStyles.Heading + "1",
            PredefinedStyles.Heading + "2",
            PredefinedStyles.Heading + "3",
            PredefinedStyles.Heading + "4",
            PredefinedStyles.Heading + "5",
            PredefinedStyles.Heading + "6",
            PredefinedStyles.Hyperlink,
            PredefinedStyles.IntenseQuote,
            PredefinedStyles.ListParagraph,
            PredefinedStyles.Quote,
            PredefinedStyles.QuoteChar,
            PredefinedStyles.TableGrid,
            PredefinedStyles.Paragraph
        ];
        this.mainPart = mainPart;
    }

    //____________________________________________________________________
    //

    /// <summary>
    /// Preload the styles in the document to match localized style name.
    /// </summary>
    internal void PrepareStyles(MainDocumentPart mainPart)
    {
        if (mainPart.StyleDefinitionsPart == null) return;

        Styles? styles = mainPart.StyleDefinitionsPart.Styles;
        if (styles == null) return;

        foreach (var s in styles.Elements<Style>())
        {
            if (s.StyleId == null)
                continue;

            if (s.StyleName != null)
            {
                string? name = s.StyleName!.Val?.Value;
                if (name != null && name != s.StyleId) knownStyles[name] = s;
            }

            knownStyles.TryAdd(s.StyleId!, s);
        }
    }

    internal ParagraphStyleId? GetParagraphStyle(string name)
    {
        var style = GetStyle(name, StyleValues.Paragraph);
        return style is null? null : new ParagraphStyleId() { Val = style };
    }
    internal RunStyle? GetRunStyle(string name)
    {
        var style = GetStyle(name, StyleValues.Character);
        return style is null? null : new RunStyle() { Val = style };
    }
    internal TableStyle? GetTableStyle(string name)
    {
        var style = GetStyle(name, StyleValues.Table);
        return style is null? null : new TableStyle() { Val = style };
    }

    /// <summary>
    /// Helper method to obtain the StyleId of a named style (invariant or localized name).
    /// </summary>
    /// <param name="name">The name of the style to look for.</param>
    /// <param name="styleType">True to obtain the character version of the given style.</param>
    /// <param name="ignoreCase">Indicate whether the search should be performed with the case-sensitive flag or not.</param>
    /// <returns>If not found, returns the given name argument.</returns>
    internal string? GetStyle(string name, StyleValues styleType, bool ignoreCase = false)
    {
        Style? style;

        // OpenXml is case-sensitive but CSS is not.
        // We will try to find the styles another time with case-insensitive:
        if (ignoreCase)
        {
            if (!knownStyles.TryGetValueIgnoreCase(name, styleType, out style))
            {
                if (StyleMissing != null)
                {
                    StyleMissing(this, new StyleEventArgs(name, styleType));
                    if (knownStyles.TryGetValueIgnoreCase(name, styleType, out style))
                        return style?.StyleId;
                }
                return null; // null means we ignore this style (css class)
            }

            return style!.StyleId;
        }
        else
        {
            if (!knownStyles.TryGetValue(name, out style))
            {
                if (lazyPredefinedStyles.Contains(name))
                {
                    string? xml = PredefinedStyles.GetOuterXml(name);
                    if (xml != null)
                        this.AddStyle(name, style = new Style(xml));
                }

                if (style is null)
                {
                    StyleMissing?.Invoke(this, new StyleEventArgs(name, styleType));
                    return name;
                }
            }

            if (styleType == StyleValues.Character && !StyleValues.Character.Equals(style!.Type!))
            {
                LinkedStyle? linkStyle = style!.GetFirstChild<LinkedStyle>();
                if (linkStyle != null) return linkStyle.Val;
            }
            return style.StyleId;
        }
    }

    /// <summary>
    /// Add a new style inside the document and refresh the style cache.
    /// </summary>
    public void AddStyle(Style style)
    {
        if (style is null) throw new ArgumentNullException(nameof(style));
        if (style.StyleId == null && style.StyleName?.Val?.HasValue != true)
            throw new ArgumentNullException($"{nameof(style)}.{nameof(style.StyleId)}");
        if (style.Type is null)
            throw new ArgumentNullException($"{nameof(style)}.{nameof(style.Type)}");

        if (style.StyleName?.Val?.HasValue != true)
            style.StyleName = new() { Val = style.StyleId };
        else if (style.StyleId?.HasValue != true)
            style.StyleId = style.StyleName.Val;

        AddStyle(style.StyleId!.Value!, style);
    }

    /// <summary>
    /// Add a new style inside the document and refresh the style cache.
    /// </summary>
    internal void AddStyle(string name, Style style)
    {
        if (name is null) throw new ArgumentNullException(nameof(name));
        if (style is null) throw new ArgumentNullException(nameof(style));
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Name cannot be empty", nameof(name));

        if (knownStyles.ContainsKey(name))
            return;

        knownStyles[name] = style;
        if (mainPart.StyleDefinitionsPart == null)
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();

        if (style.StyleName?.Val?.HasValue != true)
            style.StyleName = new() { Val = name };

        mainPart.StyleDefinitionsPart!.Styles!.Append(style);
    }

    //____________________________________________________________________
    //

    /// <summary>
    /// Contains the default styles for new OpenXML elements
    /// </summary>
    public DefaultStyles DefaultStyles
    {
        get => defaultStyles ??= new DefaultStyles();
    }

    /// <summary>
    /// Gets or sets the beginning and ending characters used in the &lt;q&gt; tag.
    /// </summary>
    public QuoteChars QuoteCharacters { get; set; } = QuoteChars.IE;
}
