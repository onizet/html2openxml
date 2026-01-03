using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests
{
    public partial class TableTests
    {
        [Test(Description = "Respect the colgroup widths")]
        public void FixedLayout_WithCol_ReturnsFixedWidth()
        {
            var elements = converter.Parse(@"<table style='table-layout:fixed'>
                <colgroup><col width='15%'><col width='85%'></colgroup>
                <tbody>
                    <tr>
                        <td style='width:120px;'>Cell 1</td>
                        <td>Cell 2</td>
                    </tr>
                </tbody>
                </table>");

            Assert.That(elements, Has.Count.EqualTo(1));
            var columns = elements[0].GetFirstChild<TableGrid>()?.Elements<GridColumn>();
            Assert.That(columns, Is.Not.Null);
            using (Assert.EnterMultipleScope())
            {
                Assert.That(columns.Count(), Is.EqualTo(2));
                Assert.That(columns.First().Width?.Value, Is.Not.EqualTo("1269"));
                Assert.That(columns.Last().Width?.Value, Is.EqualTo("8179"));
            }

            var cells = elements[0].GetFirstChild<TableRow>()?.Elements<TableCell>();
            Assert.That(cells, Is.Not.Null);
            using (Assert.EnterMultipleScope())
            {
                Assert.That(cells.Count(), Is.EqualTo(2));
                Assert.That(cells.First().TableCellProperties?.TableCellWidth, Is.Null, "Inline cell style is ignored");
            }
        }

        [Test(Description = "Colgroup doesn't contain the real number of cells")]
        public void FixedLayout_WithWrongCol_ReturnsFixedWidth()
        {
            var elements = converter.Parse(@"<table style='table-layout:fixed'>
                <colgroup><col width='30px'><col width='120px'></colgroup>
                <tbody>
                    <tr>
                        <td style='width:120px;'>Cell 1</td>
                        <td>Cell 2</td>
                        <td style='width:60px;'>Cell 3</td>
                    </tr>
                </tbody>
                </table>");

            Assert.That(elements, Has.Count.EqualTo(1));
            var columns = elements[0].GetFirstChild<TableGrid>()?.Elements<GridColumn>();
            Assert.That(columns, Is.Not.Null);
            using (Assert.EnterMultipleScope())
            {
                Assert.That(columns.Count(), Is.EqualTo(3));
                Assert.That(columns.First().Width?.Value, Is.Not.EqualTo("1269"));
                Assert.That(columns.ElementAt(1).Width?.Value, Is.EqualTo("8179"));
            }

            var cells = elements[0].GetFirstChild<TableRow>()?.Elements<TableCell>();
            Assert.That(cells, Is.Not.Null);
            using (Assert.EnterMultipleScope())
            {
                Assert.That(cells.Count(), Is.EqualTo(2));
                Assert.That(cells.Select(c => c.TableCellProperties?.TableCellWidth), Has.All.Null, "Inline cell style is ignored");
            }
        }

        [Test(Description = "When no colgroup, consider the first row cells widths")]
        public void FixedLayout_WithNoColstyle_ReturnsFixedWidthBasedOnFirstRow()
        {
            var elements = converter.Parse(@"<table style='table-layout:fixed'>
                <tbody>
                    <tr>
                        <td style='width:120px;'>Cell 1</td>
                        <td>Cell 2</td>
                    </tr>
                    <tr>
                        <td style='width:220px;'>Cell 1</td>
                        <td>Cell 2</td>
                    </tr>
                </tbody>
                </table>");

            Assert.That(elements, Has.Count.EqualTo(1));
            var columns = elements[0].GetFirstChild<TableGrid>()?.Elements<GridColumn>();
            Assert.That(columns, Is.Not.Null);
            using (Assert.EnterMultipleScope())
            {
                Assert.That(columns.Count(), Is.EqualTo(2));
                Assert.That(columns.First().Width?.Value, Is.EqualTo("1200"));
                Assert.That(columns.Last().Width?.Value, Is.Null);
            }

            var rows = elements[0].Elements<TableRow>();
            Assert.That(rows.Count(), Is.EqualTo(2));
            using (Assert.EnterMultipleScope())
            {
                var cell1_1 = rows.First().GetFirstChild<TableCell>();
                Assert.That(cell1_1, Is.Not.Null);
                Assert.That(cell1_1?.TableCellProperties?.TableCellWidth, Is.Null, "Inline cell style is ignored");

                var cell2_1 = rows.First().GetFirstChild<TableCell>();
                Assert.That(cell2_1, Is.Not.Null);
                Assert.That(cell2_1?.TableCellProperties?.TableCellWidth, Is.Null, "Inline cell style is ignored");
            }
        }

        [Test]
        public void FixedLayout_WithNoColstyle_FirstRowWithColspan_ReturnsFixedWidthBasedOnFirstRow_SkipColspan()
        {
            var elements = converter.Parse(@"<table style='table-layout:fixed'>
                <tbody>
                    <tr>
                        <td colspan='2'>Cell 1</td>
                        <td>Cell 3</td>
                    </tr>
                    <tr>
                        <td style='width:220px;'>Cell 1</td>
                        <td>Cell 2</td>
                        <td>Cell 3</td>
                    </tr>
                </tbody>
                </table>");

            Assert.That(elements, Has.Count.EqualTo(1));
            var columns = elements[0].GetFirstChild<TableGrid>()?.Elements<GridColumn>();
            Assert.That(columns, Is.Not.Null);
            using (Assert.EnterMultipleScope())
            {
                Assert.That(columns.Count(), Is.EqualTo(2));
                Assert.That(columns.First().Width?.Value, Is.EqualTo("1200"));
                Assert.That(columns.Last().Width?.Value, Is.Null);
            }

            var rows = elements[0].Elements<TableRow>();
            Assert.That(rows.Count(), Is.EqualTo(2));
            using (Assert.EnterMultipleScope())
            {
                var cell1_1 = rows.First().GetFirstChild<TableCell>();
                Assert.That(cell1_1, Is.Not.Null);
                Assert.That(cell1_1?.TableCellProperties?.TableCellWidth, Is.Null, "Inline cell style is ignored");

                var cell2_1 = rows.First().GetFirstChild<TableCell>();
                Assert.That(cell2_1, Is.Not.Null);
                Assert.That(cell2_1?.TableCellProperties?.TableCellWidth, Is.Null, "Inline cell style is ignored");
            }
        }
    }
}