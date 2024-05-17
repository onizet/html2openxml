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
        public void ParseEmptyTable(string html)
        {
            var elements = converter.Parse(html);
            Assert.That(elements, Is.Empty);
        }

        [Test(Description = "Empty cell should generate an empty Paragraph")]
        public void ParseEmptyCell()
        {
            var elements = converter.Parse(@"<table><tr><td></td></tr></table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());

            var rows = elements[0].Elements<TableRow>();
            Assert.That(rows.Count(), Is.EqualTo(1));
            var cells = rows.First().Elements<TableCell>();
            Assert.That(cells.Count(), Is.EqualTo(1));
            Assert.That(cells.First().HasChild<Paragraph>(), Is.True);
            Assert.That(cells.First().Count(c => c is not TableCellProperties), Is.EqualTo(1));
        }

        [Test(Description = "Second row does not contains complete number of cells")]
        public void ParseRowWithNoCell()
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
        public void ParseDisorderedTableParts ()
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

        [TestCase(2, 2)]
        [TestCase(1, null)]
        [TestCase(0, null)]
        public void ParseColSpan(int colSpan, int? expectedColSpan)
        {
            var elements = converter.Parse(@$"<table>
                    <tr><th colspan=""{colSpan}"">Cell 1.1</th></tr>
                </table>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Table>());
            var rows = elements[0].Elements<TableRow>();
            Assert.That(rows.Count(), Is.EqualTo(1));
            Assert.That(rows.First().Elements<TableCell>().Count(), Is.EqualTo(1));
            var cell = rows.First().GetFirstChild<TableCell>();
            Assert.That(cell.TableCellProperties.GetFirstChild<GridSpan>()?.Val?.Value, Is.EqualTo(expectedColSpan));
        }

        [TestCase("tb-lr")]
        [TestCase("vertical-lr")]
        [TestCase("tb-rl")]
        [TestCase("vertical-rl")]
        public void ParseVerticalText(string direction)
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
        }

        [Test]
        public void ParsePreAsTable()
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
                Assert.That(cell.InnerText, Is.EqualTo(preformattedText));
                Assert.That(cell.TableCellProperties?.TableCellBorders.ChildElements.Count(), Is.EqualTo(4));
                Assert.That(cell.TableCellProperties?.TableCellBorders.ChildElements, Has.All.InstanceOf<BorderType>());
                Assert.That(cell.TableCellProperties?.TableCellBorders.Elements<BorderType>().All(b => b.Val.Value == BorderValues.Single), Is.True);
            });
        }
    }
}