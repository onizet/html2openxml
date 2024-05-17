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
using System.Globalization;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of <c>pre</c> (preformatted) element.
/// </summary>
sealed class TableExpression(IHtmlElement node) : PhrasingElementExpression(node)
{
    /// <summary>Soft limit of the number of rows supported by MS Word</summary>
    public const int MaxRows = 32767;
    private readonly IHtmlTableElement tableNode = (IHtmlTableElement) node;
    private readonly Table table = new();
    private readonly TableProperties tableProperties = new();


    /*public override void CascadeStyles(OpenXmlCompositeElement element)
    {
        throw new NotImplementedException();
    }*/

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlCompositeElement> Interpret(ParsingContext context)
    {
        ComposeStyles(context);
        TableGrid grid;
        table.AddChild(tableProperties);
        table.AddChild(grid = new TableGrid());

        var columnCount = GuessColumnsCount(tableNode);
        if (columnCount == 0)
            return [];

        grid.Append(GuessGridColumns(columnCount));

        var tableContext = context.CreateChild(this);
        foreach (var part in tableNode.AsTablePartEnumerable())
        {
            var expression = new TableSectionExpression(part, columnCount);

            foreach (var element in expression.Interpret(tableContext))
            {
                context.CascadeStyles(element);
                table.AppendChild(element);
            }
        }

        //if (table);
        //    return [];

        var results = new List<OpenXmlCompositeElement> { table };

        // Prepend or append the table caption if present
        if (tableNode.Caption != null)
        {
            var captionElements = new TableCaptionExpression(table, tableNode.Caption)
                .Interpret(context);

            if (context.Converter.TableCaptionPosition == CaptionPositionValues.Above)
                results.InsertRange(0, captionElements);
            else
                results.AddRange(captionElements);
        }

        return results;
    }

    private IEnumerable<OpenXmlElement> GuessGridColumns(int columnCount)
    {
        var columns = new List<GridColumn>(columnCount);
        for (int c = 0; c < columnCount ; c++)
        {
            columns.Add(new GridColumn());
        }
        return columns;
    }

    /// <summary>
    /// OpenXml is less tolerant than HTML and expect the precise total number of columns.
    /// </summary>
    private static int GuessColumnsCount(IHtmlTableElement tableNode)
    {
        int columnCount = 0;
        foreach(var part in tableNode.AsTablePartEnumerable())
        {
            var rowNodes = part.Rows;
            var rows = new int[rowNodes.Length];

            for(int i = 0; i < rows.Length; i++)
            {
                var row = rowNodes.ElementAt(i);
                foreach (var cell in row.Cells)
                {
                    var colSpan = cell.ColumnSpan;
                    if (colSpan == 0) colSpan = 1;
                    rows[i] += colSpan;

                    var rowSpan = cell.RowSpan;
                    if (rowSpan > 1)
                    {
                        for (int si = i; si < rowSpan && si < rows.Length; si++)
                            rows[si]++;
                    }
                }
            }

            columnCount = Math.Max(rows.Max(), columnCount);
        }

        return columnCount;
    }

    protected override void ComposeStyles (ParsingContext context)
    {
        tableProperties.TableStyle = context.DocumentStyle.GetTableStyle(context.DocumentStyle.DefaultStyles.TableStyle);

        styleAttributes = node.GetStyles();
        var unit = styleAttributes.GetUnit("width");
        if (!unit.IsValid) unit = Unit.Parse(node.GetAttribute("width"));
        if (!unit.IsValid) unit = new Unit(UnitMetric.Percent, 100);

        switch (unit.Type)
        {
            case UnitMetric.Percent:
                if (unit.Value == 100)
                {
                    // Use Auto=0 instead of Pct=auto
                    // bug reported by scarhand (https://html2openxml.codeplex.com/workitem/12494)
                    tableProperties.TableWidth = new() { Type = TableWidthUnitValues.Auto, Width = "0" };
                }
                else
                {
                    tableProperties.TableWidth = new() { Type = TableWidthUnitValues.Pct, 
                        Width = (unit.Value * 50).ToString(CultureInfo.InvariantCulture) };
                }
                break;
            case UnitMetric.Point:
            case UnitMetric.Pixel:
                tableProperties.TableWidth = new() { Type = TableWidthUnitValues.Dxa, 
                    Width = unit.ValueInDxa.ToString(CultureInfo.InvariantCulture) };
                break;
        }

        foreach (string className in node.ClassList)
        {
            var matchClassName = context.DocumentStyle.GetStyle(className, StyleValues.Table, ignoreCase: true);
            if (matchClassName != null)
            {
                tableProperties.TableStyle = new() { Val = matchClassName };
                break;
            }
        }

        var align = Converter.ToParagraphAlign(node.GetAttribute("align"));
        if (align.HasValue)
            tableProperties.TableJustification = new() { Val = align.Value.ToTableRowAlignment() };

        var spacing = Convert.ToInt16(node.GetAttribute("cellspacing"));
        if (spacing > 0)
            tableProperties.TableCellSpacing = new() {
                Type = TableWidthUnitValues.Dxa, 
                Width = new Unit(UnitMetric.Pixel, spacing).ValueInDxa.ToString(CultureInfo.InvariantCulture)
        };

        var padding = Convert.ToInt16(node.GetAttribute("cellpadding"));
        if (padding > 0)
        {
            int paddingDxa = (int) new Unit(UnitMetric.Pixel, padding).ValueInDxa;

            TableCellMarginDefault cellMargin = new() {
                TableCellLeftMargin = new() { Type = TableWidthValues.Dxa, Width = (short) paddingDxa },
                TableCellRightMargin = new() { Type = TableWidthValues.Dxa, Width = (short) paddingDxa },
                TopMargin = new() { Type = TableWidthUnitValues.Dxa, Width = paddingDxa.ToString(CultureInfo.InvariantCulture) },
                BottomMargin = new() { Type = TableWidthUnitValues.Dxa, Width = paddingDxa.ToString(CultureInfo.InvariantCulture) }
            };
            tableProperties.TableCellMarginDefault = cellMargin;
        }

        // is the border=0? If so, we remove the border regardless the style in use
        if (tableNode.Border == 0)
        {
            tableProperties.TableBorders = new TableBorders() {
                TopBorder = new TopBorder { Val = BorderValues.None },
                LeftBorder = new LeftBorder { Val = BorderValues.None },
                RightBorder = new RightBorder { Val = BorderValues.None },
                BottomBorder = new BottomBorder { Val = BorderValues.None },
                InsideHorizontalBorder = new() { Val = BorderValues.None },
                InsideVerticalBorder = new() { Val = BorderValues.None }
            };
        }
        else if (tableNode.Border >= 1)
        {
            bool handleBorders = true;
            if (tableProperties.TableStyle != null)
            {
                // check whether the style in use have borders
                var s = context.MainPart.StyleDefinitionsPart?
                    .Styles?.Elements<Style>().FirstOrDefault(e => e.StyleId == tableProperties.TableStyle.Val);
                if (s?.StyleTableProperties?.TableBorders != null) handleBorders = false;
            }

            // If the border has been specified, we display the Table Grid style which display
            // its grid lines. Otherwise the default table style hides the grid lines.
            if (handleBorders)
            {
                uint borderSize = (uint) new Unit(UnitMetric.Pixel, tableNode.Border).ValueInDxa;
                tableProperties.TableBorders = new TableBorders() {
                    TopBorder = new TopBorder { Val = BorderValues.None },
                    LeftBorder = new LeftBorder { Val = BorderValues.None },
                    RightBorder = new RightBorder { Val = BorderValues.None },
                    BottomBorder = new BottomBorder { Val = BorderValues.None },
                    InsideHorizontalBorder = new() { Val = BorderValues.Single, Size = borderSize },
                    InsideVerticalBorder = new() { Val = BorderValues.Single, Size = borderSize }
                };
            }
        }
        else
        {
            var styleBorder = styleAttributes.GetBorders();
            if (!styleBorder.IsEmpty)
            {
                var tableBorders = new TableBorders {
                    LeftBorder = Converter.ToBorder<LeftBorder>(styleBorder.Left),
                    RightBorder = Converter.ToBorder<RightBorder>(styleBorder.Right),
                    TopBorder = Converter.ToBorder<TopBorder>(styleBorder.Top),
                    BottomBorder = Converter.ToBorder<BottomBorder>(styleBorder.Bottom)
                };

                tableProperties.TableBorders = tableBorders;
            }
        }
    }
}