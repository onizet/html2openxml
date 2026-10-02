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

namespace HtmlToOpenXml;

/// <summary>
/// Contains the default OpenXml style mappings used by <see cref="HtmlConverter"/>.
/// Styles selected through these mappings provide the base formatting. CSS attributes
/// defined on HTML elements are still applied on top of the resulting Word style.
/// <para>
/// Several commonly used Word styles such as headings, hyperlinks, captions, tables,
/// footnotes and quotes can be created automatically when missing from the document.
/// Customise these mappings to use your own style names during conversion.
/// </para>
/// </summary>
public class DefaultStyles
{
    /// <summary>
    /// Default style for captions (<c>figcaption</c>, <c>caption</c>).
    /// </summary>
    /// <value>Caption</value>
    public string CaptionStyle { get; set; } = PredefinedStyles.Caption;

    /// <summary>
    /// Default style for endnote texts (<c>abbr</c>, <c>acronym</c>).
    /// </summary>
    /// <value>EndnoteText</value>
    public string EndnoteTextStyle { get; set; } = PredefinedStyles.EndnoteText;

    /// <summary>
    /// Default style for endnote references (e.g., the index number of this endnote).
    /// </summary>
    /// <value>EndnoteReference</value>
    public string EndnoteReferenceStyle { get; set; } = PredefinedStyles.EndnoteReference;

    /// <summary>
    /// Default style for new footnote texts (<c>abbr</c>, <c>acronym</c>).
    /// </summary>
    /// <value>FootnoteText</value>
    public string FootnoteTextStyle { get; set; } = PredefinedStyles.FootnoteText;

    /// <summary>
    /// Default style for new footnote references (e.g., the index number of this endnote).
    /// </summary>
    /// <value>FootnoteReference</value>
    public string FootnoteReferenceStyle { get; set; } = PredefinedStyles.FootnoteReference;

    /// <summary>
    /// Default style for headings.
    /// The converter will automatically appends the heading depth level at the end of the style name.
    /// </summary>
    /// <value>Heading</value>
    public string HeadingStyle { get; set; } = PredefinedStyles.Heading;

    /// <summary>
    /// Default style for recognized numbered headings,
    /// when <see cref="HtmlConverter.SupportsHeadingNumbering" /> is <see langword="true"/>.
    /// The converter will automatically appends the heading depth level at the end of the style name.
    ///
    /// <para>
    /// When <see cref="HtmlConverter.SupportsHeadingNumbering" /> is true, using <c>NumberedHeadingStyle</c>
    /// allows the converter to properly handle the back-end sequencing (1., 1.1, 1.2).
    /// </para>
    /// </summary>
    /// <value>Heading</value>
    public string NumberedHeadingStyle { get; set; } = PredefinedStyles.Heading;

    /// <summary>
    /// Default style for hyperlinks (<c>a</c>).
    /// </summary>
    /// <value>Hyperlink</value>
    public string HyperlinkStyle { get; set; } = PredefinedStyles.Hyperlink;

    /// <summary>
    /// Default style for list paragraphs (<c>li</c>).
    /// </summary>
    /// <value>ListParagraph</value>
    public string ListParagraphStyle { get; set; } = PredefinedStyles.ListParagraph;

    /// <summary>
    /// Default style for the <c>pre</c> when <see cref="HtmlConverter.RenderPreAsTable" /> is <see langword="true"/>.
    /// </summary>
    /// <value>TableGrid</value>
    public string PreTableStyle { get; set; } = PredefinedStyles.TableGrid;

    /// <summary>
    /// Default style for quotes (<c>quote</c>, <c>cite</c>).
    /// </summary>
    /// <value>Quote</value>
    public string QuoteStyle { get; set; } = PredefinedStyles.Quote;

    /// <summary>
    /// Default style for intense quotes (<c>blockquote</c>).
    /// </summary>
    /// <value>IntenseQuote</value>
    public string IntenseQuoteStyle { get; set; } = PredefinedStyles.IntenseQuote;

    /// <summary>
    /// Default style for tables (<c>table</c>).
    /// </summary>
    /// <value>TableGrid</value>
    public string TableStyle { get; set; } = PredefinedStyles.TableGrid;

    /// <summary>
    /// Default style for any paragraphs (<c>p</c>)
    /// in the header section (<see cref="DocumentFormat.OpenXml.Packaging.HeaderPart"/> ).
    /// </summary>
    /// <value>Header</value>
    [Obsolete("Use ParagraphHeaderStyle property for clarification")]
    public string HeaderStyle { get; set; } = PredefinedStyles.Header;

    /// <summary>
    /// Default style for any paragraphs (<c>p</c>)
    /// in the header section (<see cref="DocumentFormat.OpenXml.Packaging.HeaderPart"/> ).
    /// </summary>
    /// <value>Header</value>
    public string ParagraphHeaderStyle { get; set; } = PredefinedStyles.Header;

    /// <summary>
    /// Default style for any paragraphs (<c>p</c>)
    /// in the footer section (<see cref="DocumentFormat.OpenXml.Packaging.FooterPart"/> ).
    /// </summary>
    /// <value>Footer</value>
    [Obsolete("Use ParagraphFooterStyle property for clarification")]
    public string FooterStyle { get; set; } = PredefinedStyles.Footer;

    /// <summary>
    /// Default style for any paragraphs (<c>p</c>)
    /// in the footer section (<see cref="DocumentFormat.OpenXml.Packaging.FooterPart"/> ).
    /// </summary>
    /// <value>Footer</value>
    public string ParagraphFooterStyle { get; set; } = PredefinedStyles.Footer;

    /// <summary>
    /// Default style for body paragraph (<c>body</c> or any top level tag in HTML source).
    /// </summary>
    /// <value>Normal</value>
    [Obsolete("Use ParagraphBodyStyle property for clarification")]
    public string Paragraph { get => ParagraphBodyStyle; set => ParagraphBodyStyle = value; }

    /// <summary>
    /// Default style for paragraph (<c>p</c> or <c>body</c> or any top level tag in HTML source)
    /// in the body section.
    /// </summary>
    /// <value>Normal</value>
    public string ParagraphBodyStyle { get; set; } = PredefinedStyles.Paragraph;
}