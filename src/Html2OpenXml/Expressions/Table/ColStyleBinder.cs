using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Expressions;

sealed class ColStyleBinder(bool isFixedLayout, TableColExpression? expression = null)
{
    internal const int MaxTablePortraitWidth = 9622;
    private const int MaxTableLandscapeWidth = 12996;


    private readonly bool isFixedLayout = isFixedLayout;
    private double? percentWidth;
    private TableColExpression? expression = expression;
    public bool IsWidthDefined { get; private set; }


    /// <inheritdoc/>
    public IEnumerable<OpenXmlElement> Interpret(ParsingContext context, AngleSharp.Html.Dom.IHtmlElement colNode)
    {
        var column = new GridColumn();

        var styleAttributes = colNode.GetStyles();
        var width = styleAttributes.GetUnit("width");
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
        }

        var colSpan = Convert.ToInt32(colNode.GetAttribute(AngleSharp.Dom.AttributeNames.ColSpan));
        if (colSpan == 0)
            return [column];

        var elements = new OpenXmlElement[Math.Min(colSpan, TableExpression.MaxColumns)];
        elements[0] = column;

        for (int i = 1; i < colSpan; i++)
            elements[i] = column.CloneNode(true);

        return elements;
    }

    /// <summary>
    /// /// Apply the style properties on the provided element.
    /// </summary>
    public void CascadeStyles (OpenXmlElement element)
    {
        expression?.CascadeStyles(element);

        if (element is not TableCell cell)
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
    }

    private void DistributeCellWidths(IEnumerable<TableCell> cells)
    {
        // ignore percent width as they have priority, only distribute fixed widths
        if (!cells.Any(c => c.TableCellProperties!.TableCellWidth?.Type?.Value == TableWidthUnitValues.Dxa))
            return;

        int availableWidth = MaxTablePortraitWidth;
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
    }
}