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
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a <c>td</c> or <c>th</c> element which represent a cell in a table row.
/// </summary>
sealed class TableCellExpression(IHtmlTableCellElement node) : TableElementExpressionBase(node)
{
    private readonly IHtmlTableCellElement cellNode = node;


    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        var childElements = base.Interpret (context);

        if (!childElements.Any()) // Word requires that the cell is not empty
            childElements = [new Paragraph()];

        var cell = new TableCell (cellProperties);

        if (cellNode.ColumnSpan > 1)
        {
            cellProperties.GridSpan = new() { Val = cellNode.ColumnSpan };
        }

        if (IsValidRowSpan(cellNode.RowSpan))
        {
            cellProperties.VerticalMerge = new() { Val = MergedCellValues.Restart };
        }

        cell.Append(childElements);
        return [cell];
    }

    protected override IEnumerable<OpenXmlElement> Interpret (
        ParsingContext context, IEnumerable<AngleSharp.Dom.INode> childNodes)
    {
        return BlockElementExpression.ComposeChildren(context, childNodes, paraProperties);
    }

    protected override void ComposeStyles(ParsingContext context)
    {
        base.ComposeStyles(context);

        // Manage vertical text (only for table cell)
        string? direction = styleAttributes!["writing-mode"];
        if (direction != null)
        {
            switch (direction)
            {
                case "tb-lr":
                case "vertical-lr":
                    cellProperties.TextDirection = new() { Val = TextDirectionValues.BottomToTopLeftToRight };
                    cellProperties.TableCellVerticalAlignment = new() { Val = TableVerticalAlignmentValues.Center };
                    paraProperties.Justification = new() { Val = JustificationValues.Center };
                    break;
                case "tb-rl":
                case "vertical-rl":
                    cellProperties.TextDirection = new() { Val = TextDirectionValues.TopToBottomRightToLeft };
                    cellProperties.TableCellVerticalAlignment = new() { Val = TableVerticalAlignmentValues.Center };
                    paraProperties.Justification = new() { Val = JustificationValues.Center };
                    break;
            }
        }
    }

    /// <summary>
    /// Create a minimal TableCell to fill placeholder.
    /// </summary>
    public static TableCell CreateEmpty(params OpenXmlLeafElement[] cellProperties)
    {
        return new TableCell(new Paragraph()) {
            TableCellProperties = new TableCellProperties(cellProperties) {
                TableCellWidth = new() { Width = "0" } }
        };
    }

    internal static bool IsValidRowSpan(int rowSpan)
    {
        // 1 is the default value
        // 0 means it extends until the end of the table grouping section
        return rowSpan == 0 || rowSpan > 1;
    }
}