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
sealed class TableCellExpression(IHtmlTableCellElement node) : PhrasingElementExpression(node)
{
    private readonly IHtmlTableCellElement cellNode = node;
    private readonly TableCellProperties cellProperties = new();
    private readonly ParagraphProperties paraProperties = new();
    private int columnCount = 1;

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlCompositeElement> Interpret (ParsingContext context)
    {
        var childElements = base.Interpret (context);

        if (!childElements.Any()) // Word requires that the cell is not empty
            childElements = [new Paragraph()];

        var cell = new TableCell (cellProperties);

        if (cellNode.ColumnSpan > 1)
        {
            cellProperties.GridSpan = new() { Val = cellNode.ColumnSpan };
            columnCount = cellNode.ColumnSpan;
        }

        cell.Append(childElements);
        return [cell];
    }

    protected override IEnumerable<OpenXmlCompositeElement> Interpret (
        ParsingContext context, IEnumerable<AngleSharp.Dom.INode> childNodes)
    {
        var runs = new List<Run>();
        var flowElements = new List<OpenXmlCompositeElement>();

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
                    flowElements.Add(BlockElementExpression.CombineRuns(runs, paraProperties));
                    runs.Clear();
                }

                flowElements.Add(element);
            }
        }

        if (runs.Count > 0)
            flowElements.Add(BlockElementExpression.CombineRuns(runs, paraProperties));

        return flowElements;
    }

    protected override void ComposeStyles(ParsingContext context)
    {
        base.ComposeStyles(context);

        var valign = Converter.ToVAlign(styleAttributes!["vertical-align"]);
        if (!valign.HasValue) valign = Converter.ToVAlign(cellNode.GetAttribute("valign"));
        if (!valign.HasValue)
        {
            // in Html, table cell are vertically centered by default
            valign = TableVerticalAlignmentValues.Center;
        }

        cellProperties.TableCellVerticalAlignment = new() { Val = valign };

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
                    //runProperties..j(en.CurrentTag!, new Justification() { Val = JustificationValues.Center });
                    break;
                case "tb-rl":
                case "tb":
                case "vertical-rl":
                    cellProperties.TextDirection = new() { Val = TextDirectionValues.TopToBottomRightToLeft };
                    cellProperties.TableCellVerticalAlignment = new() { Val = TableVerticalAlignmentValues.Center };
                    //htmlStyles.Tables.BeginTagForParagraph(new Justification() { Val = JustificationValues.Center });
                    break;
            }
        }
    }

    /// <summary>
    /// Create a minimal TableCell to fill placeholder.
    /// </summary>
    public static TableCell CreateEmpty()
    {
        return new TableCell(new Paragraph(
            new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve })
        )) {
            TableCellProperties = new TableCellProperties {
                TableCellWidth = new TableCellWidth() { Width = "0" },
                VerticalMerge = new VerticalMerge() }
        };
    }

    /// <summary>
    /// The space effectively occupied by this cell.
    /// </summary>
    public int ColumnCount
    {
        get => columnCount;
    }
}