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

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a <c>tbody</c>, <c>thead</c> or <c>tfoot</c> element.
/// </summary>
sealed class TablePartExpression(IHtmlTableSectionElement node, int columCount) : TableElementExpressionBase(node)
{
    private readonly IHtmlTableSectionElement tableSectionNode = node;
    private readonly int columCount = columCount;

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret (ParsingContext context)
    {
        ComposeStyles(context);

        if (tableSectionNode.Rows.Length == 0)
            yield break;

        var childContext = context.CreateChild(this);
        // row spans scope extends until the end of the table grouping section
        var rowSpans = new RowSpanCollection();
        foreach (var row in tableSectionNode.Rows)
        {
            var expression = new TableRowExpression(row, columCount, rowSpans);

            foreach (var element in expression.Interpret(childContext))
            {
                childContext.CascadeStyles(element);
                rowSpans = expression.RowSpans;

                yield return element;
            }
        }
    }
}