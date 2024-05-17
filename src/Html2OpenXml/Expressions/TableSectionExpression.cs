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
/// Process the parsing of a <c>tbody</c>, <c>thead</c> or <c>tfoot</c> element.
/// </summary>
sealed class TableSectionExpression(IHtmlTableSectionElement node, int columCount) : PhrasingElementExpression(node)
{
    private readonly IHtmlTableSectionElement tableSectionNode = node;
    private readonly int columCount = columCount;

    /// <inheritdoc/>
    public override IEnumerable<OpenXmlCompositeElement> Interpret (ParsingContext context)
    {
        ComposeStyles(context);

        if (tableSectionNode.Rows.Length == 0)
            yield break;

        var childContext = context.CreateChild(this);
        foreach (var row in tableSectionNode.Rows)
        {
            var expression = new TableRowExpression(row, columCount);

            foreach (var element in expression.Interpret(childContext))
            {
                context.CascadeStyles(element);
                yield return element;
            }
        }
    }
}