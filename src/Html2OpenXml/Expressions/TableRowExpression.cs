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
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a <c>tr</c> element which represent a row in a table.
/// </summary>
sealed class TableRowExpression : PhrasingElementExpression
{
    private readonly IHtmlTableRowElement rowNode;
    private readonly TableRowProperties rowProperties = new();
    private readonly int columCount;
    private readonly RowSpanCollection carriedRowSpans, rowSpans = [];

    public TableRowExpression(IHtmlTableRowElement node, int columCount, RowSpanCollection carriedRowSpans)
        : base (node)
    {
        rowNode = node;
        this.columCount = columCount;
        this.carriedRowSpans = carriedRowSpans;
    }

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlCompositeElement> Interpret (ParsingContext context)
    {
        ComposeStyles(context);

        // RowSpan in html requires to skip the cell declaration on the next row,
        // whilst in OpenXml, the cell must exists with the VerticalMerge=Continue property
        var cells = new List<IHtmlTableCellElement?>(columCount);
        cells.AddRange(rowNode.Cells);
        foreach (var idx in carriedRowSpans.Columns)
        {
            cells.Insert(idx, null);
        }

        if (cells.Count == 0)
            return [];

        var childContext = context.CreateChild(this);
        var tableRow = new TableRow(rowProperties);
        int occupiedColumnSpace = 0, colIndex = 0;
        foreach (var cell in cells)
        {
            // this is the cell we have inserted ourselves for carrying over the rowSpan
            if (cell == null)
            {
                tableRow.AppendChild(TableCellExpression.CreateEmpty(
                    new VerticalMerge() { Val = MergedCellValues.Continue }
                ));

                occupiedColumnSpace += carriedRowSpans.Decrement(colIndex++);
                continue;
            }

            var expression = new TableCellExpression(cell);
            foreach (var element in expression.Interpret(childContext))
            {
                context.CascadeStyles(element);
                tableRow.AppendChild(element);
            }

            // The space effectively occupied by this cell.
            occupiedColumnSpace += Math.Max(cell.ColumnSpan, 1);

            if (TableCellExpression.IsValidRowSpan(cell.RowSpan))
            {
                rowSpans.Add(colIndex, cell.RowSpan, cell.ColumnSpan);
            }
            colIndex++;
        }

        // if the row is not complete, create empty cells
        if (occupiedColumnSpace < columCount)
        {
            tableRow.AppendChild(TableCellExpression.CreateEmpty());
        }

        rowSpans.UnionWith(carriedRowSpans);
        return [tableRow];
    }

    protected override void ComposeStyles(ParsingContext context)
    {
        base.ComposeStyles(context);

        Unit unit = styleAttributes!.GetUnit("height");
        if (!unit.IsValid) unit = Unit.Parse(rowNode.GetAttribute("height"));

        switch (unit.Type)
        {
            case UnitMetric.Point:
                rowProperties.AddChild(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = (uint) (unit.Value * 20) });
                break;
            case UnitMetric.Pixel:
                rowProperties.AddChild(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = (uint) unit.ValueInDxa });
                break;
        }
    }

    /// <summary>
    /// The carried row spans.
    /// </summary>
    public RowSpanCollection RowSpans
    {
        get => rowSpans;
    }
}