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
/// Process the parsing of <c>pre</c> (preformatted) element.
/// </summary>
sealed class PreElementExpression(IHtmlElement node) : BlockElementExpression(node)
{
    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret(ParsingContext context)
    {
        ComposeStyles(context);
        var childContext = context.CreateChild(this);
        childContext.PreserveLinebreaks = true;
        childContext.CollapseWhitespaces = false;
        var childElements = Interpret(childContext, node.ChildNodes);

        // Oftenly, <pre> tag are used to renders some code examples. They look better inside a table
        if (!context.Converter.RenderPreAsTable)
            return childElements;

        TableCell cell;
        Table preTable = new(
            new TableProperties {
                TableStyle = context.DocumentStyle.GetTableStyle(context.DocumentStyle.DefaultStyles.PreTableStyle),
                TableWidth = new() { Type = TableWidthUnitValues.Auto, Width = "0" } // 100%
            },
            new TableGrid(
                new GridColumn() { Width = "5610" }),
            new TableRow(
                cell = new TableCell {
                    // Ensure the border lines are visible (regardless of the style used)
                    TableCellProperties = new() {
                        TableCellBorders = new TableCellBorders {
                            TopBorder = new TopBorder() { Val = BorderValues.Single },
                            LeftBorder = new LeftBorder() { Val = BorderValues.Single },
                            BottomBorder = new BottomBorder() { Val = BorderValues.Single },
                            RightBorder = new RightBorder() { Val = BorderValues.Single }
                        }
                    }
                })
            );

        cell.Append(childElements);

        return [preTable];
    }
}