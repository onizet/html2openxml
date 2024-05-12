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