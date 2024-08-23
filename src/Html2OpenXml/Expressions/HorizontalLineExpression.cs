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
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a <c>hr</c> element
/// by inserting an horizontal line as it stands in many emails.
/// </summary>
sealed class HorizontalLineExpression(IHtmlElement node) : HtmlDomExpression
{
    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        var paragraph = new Paragraph();
        HtmlAttributeCollection styleAttributes;
        HtmlBorder border;

        var previousElement = node.PreviousElementSibling;
        if (previousElement != null)
        {
            // If the previous paragraph contains a bottom border or is a Table, we add some spacing between the <hr>
            // and the previous element or Word will display only the last border.
            // (see Remarks: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.bottomborder%28office.14%29.aspx)
            var addSpacing = false;

            if (previousElement is IHtmlTableElement)
            {
                addSpacing = true;
            }
            else
            {
                styleAttributes = previousElement.GetStyles();
                border = styleAttributes.GetBorders();
                if (border.Bottom.IsValid && border.Bottom.Width.ValueInDxa > 0)
                {
                    addSpacing = true;
                }
            }

            if (addSpacing)
            {
                paragraph.ParagraphProperties = new ParagraphProperties { 
                    SpacingBetweenLines = new() { Before = "240" }
                };
            }
        }

        // as this paragraph has no children, it will be deleted in RemoveEmptyParagraphs()
        // in order to kept the <hr>, we force an empty run
        paragraph.Append(new Run());

        styleAttributes = node.GetStyles();
        border = styleAttributes.GetBorders();

        // Get style from border (only top) or use Default style 
        TopBorder? hrBorderStyle;
        if (!border.IsEmpty && border.Top.IsValid)
            hrBorderStyle = new TopBorder {
                Val = border.Top.Style, 
                Color = StringValue.FromString(border.Top.Color.ToHexString()),
                Size = (uint)border.Top.Width.ValueInPoint
            };
        else
            hrBorderStyle = new TopBorder() { Val = BorderValues.Single, Size = 4U };

        paragraph.ParagraphProperties ??= new();
        paragraph.ParagraphProperties.ParagraphBorders = new ParagraphBorders {
            TopBorder = hrBorderStyle
        };
        return [paragraph];
    }
}