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
using System.Globalization;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a <c>tr</c> element which represent a row in a table.
/// </summary>
sealed class TableRowExpression : TableElementExpressionBase
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
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        ComposeStyles(context);

        // RowSpan in html requires to skip the cell declaration on the next row,
        // whilst in OpenXml, the cell must exists with the VerticalMerge=Continue property
        var cells = new List<IHtmlTableCellElement?>(columCount);
        cells.AddRange(rowNode.Cells);
        foreach (var idx in carriedRowSpans.Columns)
        {
            if (idx < cells.Count) cells.Insert(idx, null);
            else cells.Add(null);
        }

        if (cells.Count == 0)
            return [];

        var rowContext = context.CreateChild(this);
        var tableRow = new TableRow(rowProperties);
        int colIndex = 0;
        int specifiedWidthCount = 0;
        foreach (var cell in cells)
        {
            // this is the cell we have inserted ourselves for carrying over the rowSpan
            if (cell == null)
            {
                int colSpan = carriedRowSpans.Decrement(colIndex);
                var mergedCell = TableCellExpression.CreateEmpty(
                    new VerticalMerge() { Val = MergedCellValues.Continue }
                );
                if (colSpan > 1) mergedCell.TableCellProperties!.GridSpan = new() { Val = colSpan };
                tableRow.AppendChild(mergedCell);

                colIndex += colSpan;
                continue;
            }

            var expression = new TableCellExpression(cell);
            foreach (var element in expression.Interpret(rowContext))
            {
                rowContext.CascadeStyles(element);
                tableRow.AppendChild(element);

                if (element.GetFirstChild<TableCellProperties>()?.HasChild<TableCellWidth>() == true)
                    specifiedWidthCount++;
            }

            if (TableCellExpression.IsValidRowSpan(cell.RowSpan))
            {
                rowSpans.Add(colIndex, cell.RowSpan, cell.ColumnSpan);
            }

            // The space effectively occupied by this cell.
            colIndex += cell.ColumnSpan;
        }

        // if the row is not complete, create empty cells
        if (colIndex < columCount)
        {
            tableRow.AppendChild(TableCellExpression.CreateEmpty());
        }

        rowSpans.UnionWith(carriedRowSpans);

        /*if (specifiedWidthCount > 0 && specifiedWidthCount < columCount)
        {
            DistributeCellWidths(tableRow.Elements<TableCell>());
        }*/

        return [tableRow];
    }

    protected override void ComposeStyles(ParsingContext context)
    {
        base.ComposeStyles(context);

        Unit unit = styleAttributes.GetUnit("height", UnitMetric.Pixel);
        if (!unit.IsValid) unit = Unit.Parse(rowNode.GetAttribute("height").AsSpan(), UnitMetric.Pixel);

        switch (unit.Metric)
        {
            case UnitMetric.Point:
                rowProperties.AddChild(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = (uint) (unit.Value * 20) });
                break;
            case UnitMetric.Pixel:
                rowProperties.AddChild(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = (uint) unit.ValueInDxa });
                break;
        }
    }

    /*private void DistributeCellWidths(IEnumerable<TableCell> cells)
    {
        // ignore percent width as they have priority, only distribute fixed widths
        if (!cells.Any(c => c.TableCellProperties!.TableCellWidth?.Type?.Value == TableWidthUnitValues.Dxa))
            return;

        int availableWidth = TableColExpression.MaxTablePortraitWidth;
        var cellsWithoutWidths = new List<TableCell>(cells.Count());
        foreach (var cell in cells)
        {
            var cellWidth = cell.TableCellProperties!.TableCellWidth;
            if (cellWidth == null || cellWidth.Type?.Value == TableWidthUnitValues.Auto)
            {
                cellsWithoutWidths.Add(cell);
                continue;
            }

            if (cellWidth.Type?.Value == TableWidthUnitValues.Dxa && cellWidth.Width?.HasValue == true)
            {
                availableWidth -= Convert.ToInt32(cellWidth.Width.Value);
            }
        }

        var widthPerCell = (availableWidth / cellsWithoutWidths.Count).ToString(CultureInfo.InvariantCulture);
        foreach (var cell in cellsWithoutWidths)
        {
            cell.TableCellProperties!.TableCellWidth = new() {
                Type = TableWidthUnitValues.Dxa,
                Width = widthPerCell
            };
        }
    }*/

    /// <summary>
    /// The carried row spans.
    /// </summary>
    public RowSpanCollection RowSpans
    {
        get => rowSpans;
    }
}