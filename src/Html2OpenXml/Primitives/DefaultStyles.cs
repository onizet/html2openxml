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
/// Contains the default styles of Word elements
/// </summary>
public class DefaultStyles
{
    /// <summary>
    /// Default style for captions
    /// </summary>
    /// <value>Caption</value>
    public string CaptionStyle { get; set; } = PredefinedStyles.Caption;

    /// <summary>
    /// Default style for new endnote texts
    /// </summary>
    /// <value>EndnoteText</value>
    public string EndnoteTextStyle { get; set; } = PredefinedStyles.EndnoteText;

    /// <summary>
    /// Default style for new endnote references
    /// </summary>
    /// <value>EndnoteReference</value>
    public string EndnoteReferenceStyle { get; set; } = PredefinedStyles.EndnoteReference;

    /// <summary>
    /// Default style for new footnote texts
    /// </summary>
    /// <value>FootnoteText</value>
    public string FootnoteTextStyle { get; set; } = PredefinedStyles.FootnoteText;

    /// <summary>
    /// Default style for new footnote references
    /// </summary>
    /// <value>FootnoteReference</value>
    public string FootnoteReferenceStyle { get; set; } = PredefinedStyles.FootnoteReference;

    /// <summary>
    /// Default style for headings
    /// Appends the level at the end of the style name
    /// </summary>
    /// <value>Heading</value>
    public string HeadingStyle { get; set; } = PredefinedStyles.Heading;

    /// <summary>
    /// Default style for hyperlinks
    /// </summary>
    /// <value>Hyperlink</value>
    public string HyperlinkStyle { get; set; } = PredefinedStyles.Hyperlink;

    /// <summary>
    /// Default style for list paragraphs
    /// </summary>
    /// <value>ListParagraph</value>
    public string ListParagraphStyle { get; set; } = PredefinedStyles.ListParagraph;

    /// <summary>
    /// Default style for the <c>pre</c> table
    /// </summary>
    /// <value>TableGrid</value>
    public string PreTableStyle { get; set; } = PredefinedStyles.TableGrid;

    /// <summary>
    /// Default style for quotes
    /// </summary>
    /// <value>Quote</value>
    public string QuoteStyle { get; set; } = PredefinedStyles.Quote;

    /// <summary>
    /// Default style for intense quotes
    /// </summary>
    /// <value>IntenseQuote</value>
    public string IntenseQuoteStyle { get; set; } = PredefinedStyles.IntenseQuote;

    /// <summary>
    /// Default style for tables
    /// </summary>
    /// <value>TableGrid</value>
    public string TableStyle { get; set; } = PredefinedStyles.TableGrid;

    /// <summary>
    /// Default style for header paragraphs.
    /// </summary>
    /// <value>Header</value>
    public string HeaderStyle { get; set; } = PredefinedStyles.Header;

    /// <summary>
    /// Default style for footer paragraphs.
    /// </summary>
    /// <value>Footer</value>
    public string FooterStyle { get; set; } = PredefinedStyles.Footer;

    /// <summary>
    /// Default style for body paragraph.
    /// </summary>
    /// <value>Normal</value>
    public string Paragraph { get; set; } = PredefinedStyles.Paragraph;
}