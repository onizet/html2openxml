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
/// Process the parsing of a <c>tr</c> element which represent a row in a table.
/// </summary>
sealed class TableRowExpression(IHtmlTableRowElement node, int columCount) : PhrasingElementExpression(node)
{
    private readonly IHtmlTableRowElement rowNode = node;
    private readonly TableRowProperties rowProperties = new();
    private readonly int columCount = columCount;


    /// <inheritdoc/>
    public override IEnumerable<OpenXmlCompositeElement> Interpret (ParsingContext context)
    {
        ComposeStyles(context);

        if (rowNode.Cells.Length == 0)
            return [];

        var childContext = context.CreateChild(this);
        TableRow tableRow = new(rowProperties);
        int effectiveColumnCount = 0;
        foreach (var cell in rowNode.Cells)
        {
            var expression = new TableCellExpression(cell);

            foreach (var element in expression.Interpret(childContext))
            {
                context.CascadeStyles(element);
                tableRow.AppendChild(element);
            }

            effectiveColumnCount += expression.ColumnCount;
        }

        // if the row is not complete, 
        if (effectiveColumnCount < columCount)
        {
            tableRow.AppendChild(TableCellExpression.CreateEmpty());
        }

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
}