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
/// Process the parsing of the style generic table element (cell, row, section or col).
/// </summary>
abstract class TableElementExpressionBase(IHtmlElement node) : PhrasingElementExpression(node)
{
    protected readonly TableCellProperties cellProperties = new();
    protected readonly ParagraphProperties paraProperties = new();



    public override void CascadeStyles(OpenXmlElement element)
    {
        base.CascadeStyles(element);

        if (paraProperties.HasChildren && element is Paragraph p)
        {
            p.ParagraphProperties ??= new();

            var knownTags = new HashSet<string>();
            foreach (var prop in p.ParagraphProperties)
            {
                if (!knownTags.Contains(prop.LocalName))
                    knownTags.Add(prop.LocalName);
            }

            foreach (var prop in paraProperties)
            {
                if (!knownTags.Contains(prop.LocalName))
                    p.ParagraphProperties.AddChild(prop.CloneNode(true));
            }
        }


        if (cellProperties.HasChildren && element is TableCell cell)
        {
            cell.TableCellProperties ??= new();

            var knownTags = new HashSet<string>();
            foreach (var prop in cell.TableCellProperties)
            {
                if (!knownTags.Contains(prop.LocalName))
                    knownTags.Add(prop.LocalName);
            }

            foreach (var prop in cellProperties)
            {
                if (!knownTags.Contains(prop.LocalName))
                    cell.TableCellProperties.AddChild(prop.CloneNode(true));
            }
        }
    }

    protected override void ComposeStyles(ParsingContext context)
    {
        base.ComposeStyles(context);

        var valign = Converter.ToVAlign(styleAttributes!["vertical-align"]);
        if (!valign.HasValue) valign = Converter.ToVAlign(node.GetAttribute("valign"));
        if (!valign.HasValue)
        {
            // in Html, table cell are vertically centered by default
            valign = TableVerticalAlignmentValues.Center;
        }

        cellProperties.TableCellVerticalAlignment = new() { Val = valign };

        var bgcolor = styleAttributes.GetColor("background-color");
        if (bgcolor.IsEmpty) bgcolor = HtmlColor.Parse(node.GetAttribute("bgcolor"));
        if (bgcolor.IsEmpty) bgcolor = styleAttributes.GetColor("background");
        if (!bgcolor.IsEmpty)
        {
            cellProperties.Shading = new() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = bgcolor.ToHexString() };
            // we apply the bgcolor on the cell level, not the run (this is an exception)
            runProperties.Shading = null;
        }

        var halign = Converter.ToParagraphAlign(styleAttributes["text-align"]);
        if (!halign.HasValue) halign = Converter.ToParagraphAlign(node.GetAttribute("align"));
        if (halign.HasValue)
        {
            paraProperties.KeepNext = new();
            paraProperties.Justification = new() { Val = halign };
        }

        var styleBorder = styleAttributes.GetBorders();
        if (!styleBorder.IsEmpty)
        {
            var borders = new TableCellBorders {
                LeftBorder = Converter.ToBorder<LeftBorder>(styleBorder.Left),
                RightBorder = Converter.ToBorder<RightBorder>(styleBorder.Right),
                TopBorder = Converter.ToBorder<TopBorder>(styleBorder.Top),
                BottomBorder = Converter.ToBorder<BottomBorder>(styleBorder.Bottom)
            };

            cellProperties.TableCellBorders = borders;
            // we apply the borders on the cell level, not the run
            runProperties.Border = null;
        }
    }
}