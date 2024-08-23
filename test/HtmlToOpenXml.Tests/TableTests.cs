using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests for <c>table</c> or <c>pre</c>.
    /// </summary>
    [TestFixture]
    public class TableTests : HtmlConverterTestBase
    {
        [TestCase("<table><tr></tr></table>", Description = "Row with no cells")]
        [TestCase("<table></table>", Description = "No rows")]
        [TestCase("<table><tbody></tbody><thead></thead><tfoot></tfoot></table>", Description = "No rows in any parts")]
        public void EmptyTable_ShouldBeIgnored(string html)
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Is.Empty);
        }

        [Test(Description = "Empty cell should generate an empty Paragraph")]
        public void EmptyCell_ReturnsEmpty()
        {
            var elements = converter.Parse(@"<table><tr><td>Next cell is empty</td><td></td></tr></table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());

            var rows = elements[0].Elements<TableRow>();
            Assert.That(rows.Count(), Is.EqualTo(1));
            var cells = rows.First().Elements<TableCell>();
            Assert.That(cells.Count(), Is.EqualTo(2));
            Assert.That(cells.All(c => c.HasChild<Paragraph>()), Is.True);
            Assert.That(cells.Last().Count(c => c is not TableCellProperties), Is.EqualTo(1));
        }

        [Test(Description = "Empty tfoot should be ignored")]
        public void EmptyTablePart_ShouldBeIgnored()
        {
            // table parts should be reordered
            var elements = converter.Parse(@"<table>
                <tbody><tr><td>Cell 1.1</td></tr></tbody>
                <tfoot></tfoot>
            </table>");

            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements[0], Is.TypeOf(typeof(Table)));

            var rows = elements[0].Elements<TableRow>();
            Assert.That(rows.Count(), Is.EqualTo(1));
        }

        [Test(Description = "Second row does not contains complete number of cells")]
        public void RowWithNoCell_ReturnsCompletelyFilledRow()
        {
            var elements = converter.Parse(@"<table>
                <tr><td>Cell 1.1</td><td>Cell 1.2</td></tr>
                <tr><td>Cell 2.1</td></tr>
                <tr><!--no cell!--></tr>
            </table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var rows = elements[0].Elements<TableRow>();
            Assert.That(rows.Count(), Is.EqualTo(2), "Row with no cells should be skipped");
            Assert.That(rows.Select(r => r.Elements<TableCell>().Count()), 
                Has.All.EqualTo(2),
                "All rows should have the same number of cells");
        }

        [Test(Description = "Respect the order header-body-footer even if provided disordered")]
        public void DisorderedTableParts_ReturnsOrderedTable ()
        {
            // table parts should be reordered
            var elements = converter.Parse(@"<table>
                <tbody><tr><td>Body</td></tr></tbody>
                <thead><tr><td>Header</td></tr></thead>
                <tfoot><tr><td>Footer</td></tr></tfoot>
            </table>");

            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements[0], Is.TypeOf(typeof(Table)));

            var rows = elements[0].Elements<TableRow>();
            Assert.That(rows.Count(), Is.EqualTo(3));
            Assert.Multiple(() =>
            {
                Assert.That(rows.ElementAt(0).InnerText, Is.EqualTo("Header"));
                Assert.That(rows.ElementAt(1).InnerText, Is.EqualTo("Body"));
                Assert.That(rows.ElementAt(2).InnerText, Is.EqualTo("Footer"));
            });
        }

        [TestCase(2u, 2)]
        [TestCase(1u, null)]
        [TestCase(0u, null)]
        public void ColSpan_ReturnsGridSpan(uint colSpan, int? expectedColSpan)
        {
            var elements = converter.Parse(@$"<table>
                    <tr><th colspan=""{colSpan}"">Cell 1.1</th></tr>
                    <tr>{("<td>Cell</td>").Repeat(Math.Max(1, colSpan))}</tr>
                </table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var rows = elements[0].Elements<TableRow>();
            Assert.That(rows.Count(), Is.EqualTo(2));

            Assert.Multiple(() =>
            {
                Assert.That(rows.First().GetFirstChild<TableCell>()?
                    .TableCellProperties?.GetFirstChild<GridSpan>()?.Val?.Value, Is.EqualTo(expectedColSpan),
                    $"Expected GridSpan={expectedColSpan}");
                Assert.That(rows.First().Elements<TableCell>().Count(), Is.EqualTo(1),
                    "1st row should contain only 1 cell");
                Assert.That(rows.Last().Elements<TableCell>().Count(), Is.EqualTo(Math.Max(1, colSpan)),
                    $"2nd row should contains {Math.Max(1, colSpan)} cells");
            });
        }

        [Test(Description = "rowSpan=0 should extend on all rows")]
        public void RowSpanZero_ReturnsExtendedToAllRows()
        {
            var elements = converter.Parse(@"<table>
                <tbody>
                    <tr><td rowspan=""0"">Cell 1.1</td><td>Cell 1.2</td><td>Cell 1.3</td></tr>
                    <tr><td>Cell 2.2</td><td>Cell 2.3</td></tr>
                    <tr><td>Cell 3.2</td><td>Cell 3.3</td></tr>
                </tbody>
                <tfoot>
                    <tr><td>Cell 4.1</td><td>Cell 4.2</td><td>Cell 4.3</td></tr>
                </tfoot>
                </table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var rows = elements[0].Elements<TableRow>().ToArray();
            Assert.That(rows, Has.Length.EqualTo(4));
            Assert.Multiple(() =>
            {
                Assert.That(rows.Select(r => r.Elements<TableCell>().Count()),
                    Has.All.EqualTo(3),
                    "All rows should have the same number of cells");
                Assert.That(rows[0].GetFirstChild<TableCell>()?.TableCellProperties?
                    .VerticalMerge?.Val?.Value, Is.EqualTo(MergedCellValues.Restart));
                Assert.That(rows[1].GetFirstChild<TableCell>()?.TableCellProperties?
                    .VerticalMerge?.Val?.Value, Is.EqualTo(MergedCellValues.Continue));
                Assert.That(rows[2].GetFirstChild<TableCell>()?.TableCellProperties?
                    .VerticalMerge?.Val?.Value, Is.EqualTo(MergedCellValues.Continue));
                Assert.That(rows[3].GetFirstChild<TableCell>()?.TableCellProperties?
                    .VerticalMerge?.Val?.Value, Is.Null,
                    "Row on tfoot should not continue the span");
            });
        }

        [Test]
        public void RowSpan_ReturnsVerticalMerge()
        {
            var elements = converter.Parse(@"<table>
                    <tr><td>Cell 1.1</td><td>Cell 1.2</td><td>Cell 1.3</td></tr>
                    <tr><td>Cell 2.1</td><td rowspan=""2"">Cell 2.2</td><td>Cell 2.3</td></tr>
                    <tr><td>Cell 3.1</td><td>Cell 3.3</td></tr>
                </table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var rows = elements[0].Elements<TableRow>();
            Assert.That(rows.Count(), Is.EqualTo(3));
            Assert.That(rows.Select(r => r.Elements<TableCell>().Count()), 
                Has.All.EqualTo(3),
                "All rows should have the same number of cells");
            
            Assert.That(rows.ElementAt(1).Elements<TableCell>().ElementAt(1)?.TableCellProperties?.VerticalMerge?.Val?.Value, Is.EqualTo(MergedCellValues.Restart));
            Assert.That(rows.ElementAt(2).Elements<TableCell>().ElementAt(1)?.TableCellProperties?.VerticalMerge?.Val?.Value, Is.EqualTo(MergedCellValues.Continue));
        }

        [Test]
        public void RowAndColumnSpan_ReturnsVerticalMergeAndGridSpan()
        {
            var elements = converter.Parse(@"<table>
                    <tr><td rowspan=""2"" colspan=""2"">Cell 1.1</td><td>Cell 1.3</td></tr>
                    <tr><td>Cell 2.3</td></tr>
                    <tr><td>Cell 3.1</td><td>Cell 3.2</td><td>Cell 3.3</td></tr>
                </table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var rows = elements[0].Elements<TableRow>();
            Assert.That(rows.Count(), Is.EqualTo(3));
            Assert.That(rows.Take(2).Select(r => r.Elements<TableCell>().Count()), 
                Has.All.EqualTo(2),
                "1st and 2nd rows should have 2 cells");
            Assert.That(rows.Last().Elements<TableCell>().Count(), 
                Is.EqualTo(3),
                "3rd row should have 3 cells");
            Assert.That(rows.First().GetFirstChild<TableCell>()?.TableCellProperties?.GridSpan?.Val?.Value, Is.EqualTo(2));
            Assert.That(rows.First().GetFirstChild<TableCell>()?.TableCellProperties?.VerticalMerge?.Val?.Value, Is.EqualTo(MergedCellValues.Restart));

            Assert.That(rows.ElementAt(1).GetFirstChild<TableCell>()?.TableCellProperties?.GridSpan?.Val?.Value, Is.EqualTo(2));
            Assert.That(rows.ElementAt(1).GetFirstChild<TableCell>()?.TableCellProperties?.VerticalMerge?.Val?.Value, Is.EqualTo(MergedCellValues.Continue));
        }

        [TestCase("tb-lr", ExpectedResult = "btLr")]
        [TestCase("vertical-lr", ExpectedResult ="btLr")]
        [TestCase("tb-rl", ExpectedResult = "tbRl")]
        [TestCase("vertical-rl", ExpectedResult = "tbRl")]
        public string? VerticalText_ReturnsTableCellWithTextAlignment(string direction)
        {
            var elements = converter.Parse(@$"<table>
                    <tr><td style=""writing-mode:{direction}"">Cell 1.1</td></tr>
                </table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var rows = elements[0].Elements<TableRow>();
            Assert.That(rows.Count(), Is.EqualTo(1));
            Assert.That(rows.First().Elements<TableCell>().Count(), Is.EqualTo(1));
            var cell = rows.First().GetFirstChild<TableCell>();
            Assert.That(cell?.TableCellProperties?.TableCellVerticalAlignment?.Val?.Value, Is.EqualTo(TableVerticalAlignmentValues.Center));
            return cell?.TableCellProperties?.TextDirection?.Val?.InnerText;
        }

        [Test(Description = "Table padding is not supported in OpenXml, only margin")]
        public void CellPadding_ReturnsTableWithCellMargin()
        {
            var elements = converter.Parse(@$"<table cellpadding=""2"">
                    <tr><td>Cell 1.1</td></tr>
                </table>");
             Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var cellMargin = elements[0].GetFirstChild<TableProperties>()?.TableCellMarginDefault;
            Assert.That(cellMargin, Is.Not.Null);

            Assert.Multiple(() =>
            {
                Assert.That(cellMargin.TableCellLeftMargin?.Width?.Value, Is.EqualTo(29));
                Assert.That(cellMargin.TableCellRightMargin?.Width?.Value, Is.EqualTo(29));
                Assert.That(cellMargin.TopMargin?.Width?.Value, Is.EqualTo("29"));
                Assert.That(cellMargin.BottomMargin?.Width?.Value, Is.EqualTo("29"));
            });
        }

        [Test]
        public void CellSpacing_ReturnsTableCellWithSpacing()
        {
            var elements = converter.Parse(@$"<table cellspacing=""2"">
                    <tr><td>Cell 1.1</td></tr>
                </table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var cellSpacing = elements[0].GetFirstChild<TableProperties>()?.TableCellSpacing;
            Assert.That(cellSpacing?.Type?.Value, Is.EqualTo(TableWidthUnitValues.Dxa));
            Assert.That(cellSpacing?.Width?.Value, Is.EqualTo("29"));
        }

        [TestCaseSource(nameof(BorderWidthCases))]
        public void HtmlBorders_ShouldSucceed(string borderAtrribute, IEnumerable<string> expectedBorderValue, IEnumerable<uint?> expectedBorderWidth)
        {
            // we specify a style which doesn't handle borders
            converter.HtmlStyles.AddStyle(new Style {
                StyleId = "NoStyle",
                Type = StyleValues.Table
            }); 
            var elements = converter.Parse($@"<table {borderAtrribute} class='NoStyle'>
                <tr><td>Cell 1 </td></tr>
                </table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var borders = elements[0].GetFirstChild<TableProperties>()?.TableBorders;
            Assert.That(borders, Is.Not.Null);
            Assert.That(borders.HasChild<BorderType>(), Is.True);
            Assert.That(new string?[] { borders.TopBorder?.Val?.InnerText,
                borders.LeftBorder?.Val?.InnerText,
                borders.RightBorder?.Val?.InnerText,
                borders.BottomBorder?.Val?.InnerText,
                borders.InsideHorizontalBorder?.Val?.InnerText,
                borders.InsideVerticalBorder?.Val?.InnerText },
                Is.EquivalentTo(expectedBorderValue));

            if (expectedBorderWidth is null)
            {
                Assert.That(borders.Elements<BorderType>().Any(b => b.Size?.HasValue == true), Is.False);
            }
            else
            {
                Assert.That(new uint?[] { borders.TopBorder?.Size?.Value,
                    borders.LeftBorder?.Size?.Value,
                    borders.RightBorder?.Size?.Value,
                    borders.BottomBorder?.Size?.Value,
                    borders.InsideHorizontalBorder?.Size?.Value,
                    borders.InsideVerticalBorder?.Size?.Value },
                    Is.EquivalentTo(expectedBorderWidth));
            }
        }

        static readonly object[] BorderWidthCases =
        [
            // Negative border should be considered as zero
            new object[] { "border='-1'", Enumerable.Repeat("none", 6), null! },
            new object[] { "border='0'", Enumerable.Repeat("none", 6), null! },
            new object[] { "border='1'",
                new string[] { "none", "none", "none", "none", "single", "single" }, 
                new uint?[] { null, null, null, null, 14, 14 } },
            new object[] { "style='border:1px;border-bottom:3px dashed'",
                new string[] { "single", "single", "single", "dashed", null!, null! },
                new uint?[] { 6, 6, 6, 18, null, null } }
        ];

        [TestCase("above", 0, 1)]
        [TestCase("below", 1, 0)]
        public void TableCaption_ReturnsPositionedParagraph(string position, int captionPos, int tablePos)
        {
            converter.TableCaptionPosition = new (position);
            var elements = converter.Parse(@$"<table>
                    <caption>Some table caption</caption>
                    <tr><td>Cell 1.1</td></tr>
                </table>");
            Assert.That(elements, Has.Count.EqualTo(2));
            Assert.That(elements[captionPos], Is.TypeOf<Paragraph>());
            Assert.That(elements[tablePos], Is.TypeOf<Table>());
            var p = (Paragraph) elements[captionPos];
            var runs = p.Elements<Run>();
            Assert.That(runs.Count(), Is.AtLeast(4));

            Assert.Multiple(() =>{
                Assert.That(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value, Is.EqualTo(converter.HtmlStyles.DefaultStyles.CaptionStyle));
                Assert.That(runs.First().HasChild<FieldChar>(), Is.True);
                Assert.That(runs.ElementAt(1).HasChild<FieldCode>(), Is.True);
                Assert.That(runs.ElementAt(2).HasChild<FieldChar>(), Is.True);
            });
            Assert.Multiple(() =>
            {
                Assert.That(runs.First().GetFirstChild<FieldChar>()?.FieldCharType?.Value, Is.EqualTo(FieldCharValues.Begin));
                Assert.That(runs.ElementAt(1).GetFirstChild<FieldCode>()?.InnerText, Is.EqualTo("SEQ TABLE \\* ARABIC"));
                Assert.That(runs.ElementAt(2).GetFirstChild<FieldChar>()?.FieldCharType?.Value, Is.EqualTo(FieldCharValues.End));
                Assert.That(runs.Last().InnerText, Is.EqualTo(" Some table caption"));
            });
        }

        [TestCase("right", "right")]
        [TestCase("", "center")]
        public void TableCaptionAlign_ReturnsPositionedParagraph_AlignedWithTable(string alignment, string expectedAlign)
        {
            var elements = converter.Parse(@$"<table align=""center"">
                    <caption align=""{alignment}"">Some table caption</caption>
                    <tr><td>Cell 1.1</td></tr>
                </table>");

            Assert.That(elements, Has.Count.EqualTo(2));
            var caption = (Paragraph) elements[1];
            Assert.That(caption.ParagraphProperties?.Justification?.Val?.ToString(), Is.EqualTo(expectedAlign));
        }

        [Test]
        public void EmptyTableCaption_ShouldBeIgnored()
        {
            var elements = converter.Parse(@$"<table>
                    <caption></caption>
                    <tr><td>Cell 1.1</td></tr>
                </table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements[0], Is.TypeOf<Table>());
        }

        [Test]
        public void PreAsTable_ReturnsTableWithPreservedWhitespaces()
        {
            const string preformattedText = @"
              ^__^
              (oo)\_______
              (__)\       )\/\
                  ||----w |
                  ||     ||";

            converter.RenderPreAsTable = true;
            var elements = converter.Parse(@$"
<pre role='img' aria-label='ASCII COW'>
{preformattedText}</pre>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var tableProps = elements[0].GetFirstChild<TableProperties>();
            Assert.That(tableProps, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(tableProps.GetFirstChild<TableStyle>()?.Val?.Value, Is.EqualTo(converter.HtmlStyles.DefaultStyles.PreTableStyle));
                Assert.That(tableProps.GetFirstChild<TableWidth>()?.Type?.Value, Is.EqualTo(TableWidthUnitValues.Auto));
                Assert.That(tableProps.GetFirstChild<TableWidth>()?.Width?.Value, Is.EqualTo("0"));
            });

            var rows = elements[0].Elements<TableRow>();
            Assert.That(rows.Count(), Is.EqualTo(1));
            var cells = rows.First().Elements<TableCell>();
            Assert.That(cells.Count(), Is.EqualTo(1));
            var cell = cells.First();
            Assert.Multiple(() =>
            {
                Assert.That(cell.TableCellProperties?.TableCellBorders?.ChildElements.Count(), Is.EqualTo(4));
                Assert.That(cell.TableCellProperties?.TableCellBorders?.ChildElements, Has.All.InstanceOf<BorderType>());
                Assert.That(cell.TableCellProperties?.TableCellBorders?.Elements<BorderType>().All(b => b.Val?.Value == BorderValues.Single), Is.True);
            });

            var run = cell.GetFirstChild<Paragraph>()?.GetFirstChild<Run>();
            Assert.That(run, Is.Not.Null);
            Assert.Multiple(() =>
            {
                var odds = run.ChildElements.Where((item, index) => index % 2 != 0);
                Assert.That(odds, Has.All.TypeOf<Break>());
                Assert.That(run.ChildElements.ElementAt(0).InnerText, Is.EqualTo("              ^__^"));
                Assert.That(run.ChildElements.ElementAt(2).InnerText, Is.EqualTo("              (oo)\\_______"));
                Assert.That(run.ChildElements.ElementAt(4).InnerText, Is.EqualTo("              (__)\\       )\\/\\"));
                Assert.That(run.ChildElements.ElementAt(6).InnerText, Is.EqualTo("                  ||----w |"));
                Assert.That(run.ChildElements.ElementAt(8).InnerText, Is.EqualTo("                  ||     ||"));
            });
        }

        [Test]
        public void RowStyle_ReturnsRunCellsWithCascadedStyle()
        {
            var elements = converter.Parse(@$"<table>
                    <tr style='background:silver;height:120px'><td>Cell</td></tr>
                </table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());

            var row = elements[0].GetFirstChild<TableRow>();
            Assert.That(row, Is.Not.Null);
            Assert.That(row.TableRowProperties?.GetFirstChild<TableRowHeight>()?.Val?.Value, Is.EqualTo(1800));

            var cell = row.GetFirstChild<TableCell>();
            Assert.That(cell, Is.Not.Null);
            Assert.That(cell.TableCellProperties, Is.Not.Null);
            Assert.That(cell.TableCellProperties.Shading?.Fill?.Value, Is.EqualTo("C0C0C0"));

            var runProperties = cell.GetFirstChild<Paragraph>()?.GetFirstChild<Run>()?.RunProperties;
            Assert.That(runProperties?.Shading, Is.Null);
        }

        [Test]
        public void CellStyle_ReturnsRunWithCascadeStyle()
        {
            var elements = converter.Parse(@$"<table>
                    <tr><td style=""font-weight:bold""><i>Cell</i></td></tr>
                </table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var cell = elements[0].GetFirstChild<TableRow>()?.GetFirstChild<TableCell>();
            Assert.That(cell, Is.Not.Null);
            var runProperties = cell.GetFirstChild<Paragraph>()?.GetFirstChild<Run>()?.RunProperties;
            Assert.That(runProperties, Is.Not.Null);
            Assert.Multiple(() => {
                Assert.That(runProperties.Bold, Is.Not.Null);
                Assert.That(runProperties.Italic, Is.Not.Null);
            });
            Assert.Multiple(() => {
                // normally, Val should be null
                if (runProperties.Bold.Val is not null)
                    Assert.That(runProperties.Bold.Val, Is.EqualTo(true));
                if (runProperties.Italic.Val is not null)
                    Assert.That(runProperties.Italic.Val, Is.EqualTo(true));
            });
        }

        [Test]
        public void NestedTable_ReturnsTableInsideTable()
        {
            var elements = converter.Parse(@$"<table>
                    <tr><td style=""font-weight:bold"">
                        <table><tr><td>Cell</td></tr></table>
                    </td></tr>
                </table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            Assert.That(elements[0].GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count(), Is.EqualTo(1));
            var cell = elements[0].GetFirstChild<TableRow>()?.GetFirstChild<TableCell>();
            Assert.That(cell, Is.Not.Null);
            Assert.That(cell.HasChild<Table>(), Is.True);
        }

        [Test]
        public void Colstyle_ReturnsStyleAppliedOnCell()
        {
            var elements = converter.Parse(@$"<table>
                    <colgroup>
                        <col style=""width:100px""/>
                        <col style=""width:50px;border:3px double #000000""/>
                    </colgroup>
                    <tr><td>Cell 1.1</td><td>Cell 1.2</td></tr>
                </table>");

            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var columns = elements[0].GetFirstChild<TableGrid>()?.Elements<GridColumn>();
            Assert.That(columns, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(columns.Count(), Is.EqualTo(2));
                //Assert.That(columns.First().Width?.Value, Is.EqualTo("1500"));
                //Assert.That(columns.Last().Width?.Value, Is.EqualTo("750"));
            });

            var cells = elements[0].GetFirstChild<TableRow>()?.Elements<TableCell>();
            Assert.That(cells, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(cells.Count(), Is.EqualTo(2));
                Assert.That(cells.First().TableCellProperties?.TableCellBorders, Is.Null);
                Assert.That(cells.Last().TableCellProperties?.TableCellBorders, Is.Not.Null);
            });
        }

        [Test]
        public void ColWithSpan_ReturnsStyleAppliedOnMultipleCells()
        {
            var elements = converter.Parse(@$"<table>
                    <colgroup>
                        <col style=""width:100px"" span=""2"" align=""right"" />
                        <col style=""width:50px""/>
                    </colgroup>
                    <tr><td>Cell 1.1</td><td>Cell 1.2</td><td>Cell 1.3</td></tr>
                </table>");

            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var columns = elements[0].GetFirstChild<TableGrid>()?.Elements<GridColumn>();
            Assert.That(columns, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(columns.Count(), Is.EqualTo(3));
                //Assert.That(columns.First().Width?.Value, Is.EqualTo("1500"));
                //Assert.That(columns.ElementAt(1).Width?.Value, Is.EqualTo("1500"));
                //Assert.That(columns.Last().Width?.Value, Is.EqualTo("750"));
            });

            var cells = elements[0].GetFirstChild<TableRow>()?.Elements<TableCell>();
            Assert.That(cells, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(cells.Count(), Is.EqualTo(3));
                Assert.That(cells.First().GetFirstChild<Paragraph>()?.ParagraphProperties?.Justification?.Val?.Value, Is.EqualTo(JustificationValues.Right));
                Assert.That(cells.ElementAt(1).GetFirstChild<Paragraph>()?.ParagraphProperties?.Justification?.Val?.Value, Is.EqualTo(JustificationValues.Right));
                Assert.That(cells.Last().GetFirstChild<Paragraph>()?.ParagraphProperties?.Justification?.Val, Is.Null);
            });
        }

        [Test(Description = "Table row contains more cell than specified col")]
        public void IncompleteColStyle_ShouldNotThrow()
        {
            Assert.DoesNotThrow(() => converter.Parse(@$"<table>
                    <colgroup>
                        <col style=""width:100px""/>
                    </colgroup>
                    <tr><td>Cell 1.1</td><td>Cell 1.2</td></tr>
                </table>"));
        }

        [Test(Description = "Cell with multiple runs")]
        public void CellWithMultipleDivs_ReturnsOneCellWithMultipleParagraphs()
        {
            var elements = converter.Parse(@$"<table>
                    <tr><td>Cell <div><b>1.1</b></div></td></tr>
                </table>");

            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var cells = elements[0].GetFirstChild<TableRow>()?.Elements<TableCell>();
            Assert.That(cells?.Count(), Is.EqualTo(1));
            Assert.That(cells.First().Elements<Paragraph>().Count(), Is.EqualTo(2));
        }

        [Test(Description = "Prevent Word to merge two consecutive tables")]
        public void ConsecutiveTables_ReturnsTablesWithPlaceholderParagraphInBetween()
        {
            var elements = converter.Parse(@"
                <table><tr><td>Table 1</td></tr></table>
                <table><tr><td>Table 2</td></tr></table>");
            Assert.That(elements, Has.Count.EqualTo(3));
            Assert.Multiple(() =>
            {
                Assert.That(elements[0], Is.TypeOf<Table>());
                Assert.That(elements[1], Is.TypeOf<Paragraph>());
                Assert.That(elements[2], Is.TypeOf<Table>());
            });
        }

        [Test(Description = "Insert the empty paragraph on right location, based on combinated rowSpan and colSpan #59")]
        public void DoubleRowSpanSameLine_ShouldSucceed()
        {
            var elements = converter.Parse(@"<table>
                    <tr>
                        <td rowspan=""2"" colspan=""2"">Cell 1.1</td>
                        <td>Cell 1.2</td>
                        <td rowspan=""2"">Cell 1.3</td>
                    </tr>
                    <tr>
                        <td>Cell 2.2</td>
                    </tr>
                </table>");

            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            Assert.That(elements[0].GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count(), Is.EqualTo(4));

            var rows = elements[0].Elements<TableRow>();
            Assert.That(rows.Count(), Is.EqualTo(2));
            Assert.That(rows.Select(r => r.Elements<TableCell>().Count()), 
                Has.All.EqualTo(3),
                "All should have 3 cells");
            Assert.That(rows.First().GetFirstChild<TableCell>()?.TableCellProperties?.GridSpan?.Val?.Value, Is.EqualTo(2));
            Assert.That(rows.First().GetFirstChild<TableCell>()?.TableCellProperties?.VerticalMerge?.Val?.Value, Is.EqualTo(MergedCellValues.Restart));
            Assert.That(rows.First().GetLastChild<TableCell>()?.TableCellProperties?.VerticalMerge?.Val?.Value, Is.EqualTo(MergedCellValues.Restart));

            Assert.That(rows.ElementAt(1).GetFirstChild<TableCell>()?.TableCellProperties?.GridSpan?.Val?.Value, Is.EqualTo(2));
            Assert.That(rows.ElementAt(1).GetFirstChild<TableCell>()?.TableCellProperties?.VerticalMerge?.Val?.Value, Is.EqualTo(MergedCellValues.Continue));
            Assert.That(rows.ElementAt(1).GetLastChild<TableCell>()?.TableCellProperties?.VerticalMerge?.Val?.Value, Is.EqualTo(MergedCellValues.Continue));
        }

        [Test(Description = "Cell borders should not be applied on inner runs #156")]
        public void CellBorders_ShouldNotPropagate_OnRuns()
        {
            var elements = converter.Parse(@"<table>
                <tr>
                    <td style='width: 120px; border: 1px solid #ababab'>Placeholder</td>
                </tr>
            </table>");

            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var cells = elements[0].GetFirstChild<TableRow>()?.Elements<TableCell>();
            Assert.That(cells?.Count(), Is.EqualTo(1));
            Assert.That(cells.First().TableCellProperties?.TableCellBorders, Is.Not.Null);
            var paragraphs = cells.First().Elements<Paragraph>();
            Assert.That(paragraphs.Count(), Is.EqualTo(1));
            var runs = paragraphs.First().Elements<Run>();
            Assert.That(runs.Count(), Is.EqualTo(1));
            Assert.That(runs.First().RunProperties?.Border, Is.Null);
        }
    }
}