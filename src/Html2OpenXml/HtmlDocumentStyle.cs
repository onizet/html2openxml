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

    private readonly RunStyleCollection runStyle;
    private TableStyleCollection tableStyle;
    private readonly ParagraphStyleCollection paraStyle;
    private NumberingListStyleCollection listStyle;
    private readonly MainDocumentPart mainPart;
    private readonly OpenXmlDocumentStyleCollection knownStyles;
    private readonly ISet<string> lazyPredefinedStyles;

    private DefaultStyles? defaultStyles;
    

    internal WordDocumentStyle(MainDocumentPart mainPart)
    {
        knownStyles = PrepareStyles(mainPart);
        lazyPredefinedStyles = new HashSet<string>() { 
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
            PredefinedStyles.TableGrid
        };
        tableStyle = new TableStyleCollection(this);
        runStyle = new RunStyleCollection(this);
        paraStyle = new ParagraphStyleCollection(this);
        this.mainPart = mainPart;
    }

    //____________________________________________________________________
    //

    /// <summary>
    /// Preload the styles in the document to match localized style name.
    /// </summary>
    internal OpenXmlDocumentStyleCollection PrepareStyles(MainDocumentPart mainPart)
    {
        var knownStyles = new OpenXmlDocumentStyleCollection();
        if (mainPart.StyleDefinitionsPart == null) return knownStyles;

        Styles? styles = mainPart.StyleDefinitionsPart.Styles;
        if (styles == null) return knownStyles;

        foreach (var s in styles.Elements<Style>())
        {
            if (s.StyleId == null)
                continue;

            if (s.StyleName != null)
            {
                string? name = s.StyleName!.Val?.Value;
                if (name != null && name != s.StyleId) knownStyles[name] = s;
            }

            knownStyles.Add(s.StyleId!, s);
        }
        return knownStyles;
    }

    internal ParagraphStyleId GetParagraphStyle(string name)
    {
        return new ParagraphStyleId() { Val = GetStyle(name, StyleValues.Paragraph) };
    }
    internal RunStyle GetRunStyle(string name)
    {
        return new RunStyle { Val = GetStyle(name, StyleValues.Character) };
    }
    internal TableStyle GetTableStyle(string name)
    {
        return new TableStyle { Val = GetStyle(name, StyleValues.Table) };
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
                    StyleMissing(this, new StyleEventArgs(name, mainPart.StyleDefinitionsPart!, styleType));
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
                    StyleMissing?.Invoke(this, new StyleEventArgs(name, mainPart.StyleDefinitionsPart!, styleType));
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
    private void AddStyle(string name, Style style)
    {
        knownStyles[name] = style;
        if (mainPart.StyleDefinitionsPart == null)
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
        mainPart.StyleDefinitionsPart!.Styles!.Append(style);
    }

    //____________________________________________________________________
    //

    [Obsolete]
    internal RunStyleCollection Runs
    {
        [System.Diagnostics.DebuggerHidden()]
        get { return runStyle; }
    }
    internal TableStyleCollection Tables
    {
        [System.Diagnostics.DebuggerHidden()]
        get { return tableStyle; }
    }
    [Obsolete]
    internal ParagraphStyleCollection Paragraph
    {
        [System.Diagnostics.DebuggerHidden()]
        get { return paraStyle; }
    }
    internal NumberingListStyleCollection NumberingList
    {
        // use lazy loading to avoid injecting NumberListDefinition if not required
        [System.Diagnostics.DebuggerHidden()]
        get { return listStyle ?? (listStyle = new NumberingListStyleCollection(mainPart)); }
    }

    /// <summary>
    /// Contains the default styles for new OpenXML elements
    /// </summary>
    public DefaultStyles DefaultStyles
    {
        get => defaultStyles ??= new DefaultStyles();
        set => defaultStyles = value;
    }

    /// <summary>
    /// Gets or sets the beginning and ending characters used in the &lt;q&gt; tag.
    /// </summary>
    public QuoteChars QuoteCharacters { get; set; } = QuoteChars.IE;
}
