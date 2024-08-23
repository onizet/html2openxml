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
sealed class TableExpression(IHtmlTableElement node) : PhrasingElementExpression(node)
{
    /// <summary>MS Word has this hard-limit.</summary>
    internal const int MaxColumns = short.MaxValue;
    private readonly IHtmlTableElement tableNode = node;
    private readonly Table table = new();
    private readonly TableProperties tableProperties = new();
    private TableColExpression[]? colStyleExpressions;
    private int colIndex;


    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret(ParsingContext context)
    {
        ComposeStyles(context);
        TableGrid grid;
        table.AddChild(tableProperties);
        table.AddChild(grid = new TableGrid());

        var columnCount = GuessColumnsCount(tableNode);
        if (columnCount == 0)
            return [];

        grid.Append(InterpretGridColumns(context, columnCount));

        var tableContext = context.CreateChild(this);
        foreach (var part in tableNode.AsTablePartEnumerable())
        {
            var expression = new TablePartExpression(part, columnCount);

            foreach (var element in expression.Interpret(tableContext))
            {
                context.CascadeStyles(element);
                table.AppendChild(element);
            }
        }

        var results = new List<OpenXmlElement> { table };

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

    /// <summary>
    /// Parse the <c>col</c> tags, defining some column styles.
    /// </summary>
    private IEnumerable<GridColumn> InterpretGridColumns(ParsingContext context, int columnCount)
    {
        var columns = new List<GridColumn>(columnCount);
        var colStyleExpressions = new List<TableColExpression>(columnCount);

        var colgroup = tableNode.Children.FirstOrDefault(n => n.LocalName == "colgroup");
        // if colgroup tag is not found, maybe the table was misformed and they stand below the root level
        colgroup ??= tableNode;

        foreach (var col in colgroup.Children
            .Where(n => n.LocalName == "col").Cast<IHtmlTableColumnElement>())
        {
            var expression = new TableColExpression(col);
            foreach (var child in expression.Interpret(context).Cast<GridColumn>())
            {
                columns.Add(child);
                colStyleExpressions.Add(expression);
            }
        }

        for (int c = columns.Count; c < columnCount ; c++)
        {
            columns.Add(new GridColumn());
        }

        if (colStyleExpressions.Count > 0)
        {
            this.colStyleExpressions = [.. colStyleExpressions];
        }

        return columns;
    }

    public override void CascadeStyles(OpenXmlElement element)
    {
        base.CascadeStyles(element);

        if (colStyleExpressions != null)
        {
            if (element is TableRow)
            {
                colIndex = 0;
            }

            if (colIndex < colStyleExpressions.Length)
                colStyleExpressions![colIndex].CascadeStyles(element);

            if (element is TableCell)
            {
                colIndex++;
            }
        }
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
                foreach (var cell in rowNodes.ElementAt(i).Cells)
                {
                    var colSpan = Math.Max(1, cell.ColumnSpan);
                    for (int r = i; r < i + cell.RowSpan; r++)
                    {
                        rows[r] += colSpan;
                    }
                }
            }

            if (rows.Any())
                columnCount = Math.Max(rows.Max(), columnCount);
        }

        return Math.Min(columnCount, MaxColumns);
    }

    protected override void ComposeStyles (ParsingContext context)
    {
        tableProperties.TableStyle = context.DocumentStyle.GetTableStyle(context.DocumentStyle.DefaultStyles.TableStyle);

        styleAttributes = tableNode.GetStyles();
        var width = styleAttributes.GetUnit("width", UnitMetric.Pixel);
        if (!width.IsValid) width = Unit.Parse(tableNode.GetAttribute("width"), UnitMetric.Pixel);
        if (!width.IsValid) width = new Unit(UnitMetric.Percent, 100);

        switch (width.Type)
        {
            case UnitMetric.Percent:
                if (width.Value == 100)
                {
                    // Use Auto=0 instead of Pct=auto
                    // bug reported by scarhand (https://html2openxml.codeplex.com/workitem/12494)
                    tableProperties.TableWidth = new() { Type = TableWidthUnitValues.Auto, Width = "0" };
                }
                else
                {
                    tableProperties.TableWidth = new() { Type = TableWidthUnitValues.Pct, 
                        Width = (width.Value * 50).ToString(CultureInfo.InvariantCulture) };
                }
                break;
            case UnitMetric.Point:
            case UnitMetric.Pixel:
                tableProperties.TableWidth = new() { Type = TableWidthUnitValues.Dxa, 
                    Width = width.ValueInDxa.ToString(CultureInfo.InvariantCulture) };
                break;
        }

        foreach (string className in tableNode.ClassList)
        {
            var matchClassName = context.DocumentStyle.GetStyle(className, StyleValues.Table, ignoreCase: true);
            if (matchClassName != null)
            {
                tableProperties.TableStyle = new() { Val = matchClassName };
                break;
            }
        }

        var align = Converter.ToParagraphAlign(tableNode.GetAttribute("align"));
        if (align.HasValue)
            tableProperties.TableJustification = new() { Val = align.Value.ToTableRowAlignment() };

        var dir = tableNode.GetTextDirection();
        if (dir.HasValue)
            tableProperties.BiDiVisual = new() { 
                Val = dir == AngleSharp.Dom.DirectionMode.Rtl? OnOffOnlyValues.On : OnOffOnlyValues.Off
            };

        var spacing = Convert.ToInt16(tableNode.GetAttribute("cellspacing"));
        if (spacing > 0)
            tableProperties.TableCellSpacing = new() {
                Type = TableWidthUnitValues.Dxa, 
                Width = new Unit(UnitMetric.Pixel, spacing).ValueInDxa.ToString(CultureInfo.InvariantCulture)
        };

        var padding = Convert.ToInt16(tableNode.GetAttribute("cellpadding"));
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

        var styleBorder = styleAttributes.GetBorders();

        if (!styleBorder.IsEmpty)
        {
            var tableBorders = new TableBorders {
                TopBorder = Converter.ToBorder<TopBorder>(styleBorder.Top),
                LeftBorder = Converter.ToBorder<LeftBorder>(styleBorder.Left),
                RightBorder = Converter.ToBorder<RightBorder>(styleBorder.Right),
                BottomBorder = Converter.ToBorder<BottomBorder>(styleBorder.Bottom)
            };

            tableProperties.TableBorders = tableBorders;
        }
        // is the border=0? If so, we remove the border regardless the style in use
        else if (tableNode.Border == 0)
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
    }
}