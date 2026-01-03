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
//using System.Globalization;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

/// <summary>
/// Process the parsing of a <c>col</c> element.
/// </summary>
sealed class TableColExpression(IHtmlTableColumnElement node) : TableElementExpressionBase(node)
{
    internal const int MaxTablePortraitWidth = 9622;
    /// <summary>
    /// private const int MaxTableLandscapeWidth = 12996;
    /// </summary>
    private readonly IHtmlTableColumnElement colNode = node;
   // private readonly bool isFixedLayout;
    //protected double? percentWidth;
    //public bool IsWidthDefined { get; private set; }



    /// <inheritdoc/>
    public override IEnumerable<OpenXmlElement> Interpret(ParsingContext context)
    {
        ComposeStyles(context);

        var column = new GridColumn();
        /*var width = styleAttributes.GetUnit("width");
        if (!width.IsValid) width = Unit.Parse(colNode!.GetAttribute("width"));

        if (width.IsValid)
        {
            IsWidthDefined = true;
            if (width.IsFixed)
            {
                // This value is specified in twentieths of a point.
                // If this attribute is omitted, then the last saved width of the grid column is assumed to be zero.
                column.Width = Math.Round(width.ValueInPoint * 20).ToString(CultureInfo.InvariantCulture);
            }
            else if (width.Metric == UnitMetric.Percent)
            {
                var maxWidth = context.IsLandscape ? MaxTableLandscapeWidth : MaxTablePortraitWidth;
                percentWidth = Math.Max(0, Math.Min(100, width.Value));
                column.Width = Math.Ceiling(maxWidth / 100d * percentWidth.Value).ToString(CultureInfo.InvariantCulture);
            }
        }*/

        if (colNode.Span == 0)
            return [column];

        var elements = new OpenXmlElement[Math.Min(colNode.Span, TableExpression.MaxColumns)];
        elements[0] = column;

        for (int i = 1; i < colNode.Span; i++)
            elements[i] = column.CloneNode(true);

        return elements;
    }

    /*public override void CascadeStyles(OpenXmlElement element)
    {
        base.CascadeStyles(element);

        f (element is not TableCell cell)
            return;

        if (isFixedLayout)
        {
            // in fixed layout, we ignore any inline width style
            cell.TableCellProperties!.TableCellWidth = null;
        }
        else if (percentWidth.HasValue && cell.TableCellProperties?.TableCellWidth is null)
        {
            cell.TableCellProperties!.TableCellWidth = new() {
                Type = TableWidthUnitValues.Pct,
                Width = ((int)(percentWidth.Value * 50)).ToString(CultureInfo.InvariantCulture)
            };
        }
    }*/
}