using NUnit.Framework;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HtmlToOpenXml.Tests
{
    /// <summary>
    /// Tests on <c>div</c> and other block elements.
    /// </summary>
    [TestFixture]
    public class DivTests : HtmlConverterTestBase
    {
        [Test]
        public void StyleAttribute_WithMultipleValues_ShouldBeAllApplied()
        {
            var elements = converter.Parse(@"<div style='text-indent:1em;border:1px dotted red;text-align:center;line-height:2.5;margin-right:1em'>Lorem</div>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            var p = (Paragraph) elements[0];
            Assert.That(p.ParagraphProperties, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(p.ParagraphProperties.Indentation?.FirstLine?.HasValue, Is.True);
                Assert.That(p.ParagraphProperties.ParagraphBorders, Is.Not.Null);
                Assert.That(p.ParagraphProperties.Justification?.Val?.Value, Is.EqualTo(JustificationValues.Center));
                Assert.That(p.ParagraphProperties.SpacingBetweenLines?.Line?.Value, Is.EqualTo("600"));
                Assert.That(p.ParagraphProperties.Indentation?.Right?.Value, Is.EqualTo("239"));
            });

            var borders = p.ParagraphProperties?.ParagraphBorders?.Elements<BorderType>();
            Assert.That(borders, Is.Not.Null);
            Assert.Multiple(() =>
            {
                Assert.That(borders.Count(), Is.EqualTo(4));
                Assert.That(borders.Select(b => b.Color?.Value), Has.All.EqualTo("FF0000"));
                Assert.That(borders.Select(b => b.Val?.Value), Has.All.EqualTo(BorderValues.Dotted));
            });
        }

        [Test]
        public void PageBreakBefore_ReturnsOneParagraphThenTwo()
        {
            var elements = converter.Parse(@"Lorem
                <div style='page-break-before:always'>Placeholder</div>
                Ipsum");
            Assert.That(elements, Has.Count.EqualTo(3));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            Assert.That(elements[0].ChildElements, Has.Count.EqualTo(1));
            Assert.Multiple(() =>
            {
                Assert.That(elements[0].ChildElements, Has.All.TypeOf<Run>());
                Assert.That(elements[0].InnerText, Is.EqualTo("Lorem"));
                Assert.That(elements[1].ChildElements, Has.Count.EqualTo(3));
                Assert.That(elements[2].ChildElements, Has.All.TypeOf<Run>());
                Assert.That(elements[2].InnerText, Is.EqualTo("Ipsum"));
                Assert.That(elements[2].ChildElements, Has.Count.EqualTo(1));
            });
            Assert.Multiple(() =>
            {
                Assert.That(elements[1].ChildElements, Has.All.TypeOf<Run>());
                Assert.That(elements[1].ChildElements[0].HasChild<Break>(), Is.True);
                Assert.That(elements[1].ChildElements[1].HasChild<LastRenderedPageBreak>(), Is.True);
                Assert.That(elements[1].ChildElements[2].InnerText, Is.EqualTo("Placeholder"));
                Assert.That(elements[1].InnerText, Is.EqualTo("Placeholder"));
            });
        }

        [Test]
        public void PageBreakAfter_ReturnsTwoParagraphsThenOne()
        {
            var elements = converter.Parse(@"Lorem
                <div style='page-break-after:always'>Placeholder</div>
                Ipsum");
            Assert.That(elements, Has.Count.EqualTo(3));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            Assert.That(elements[0].ChildElements, Has.Count.EqualTo(1));
            Assert.Multiple(() =>
            {
                Assert.That(elements[0].ChildElements, Has.All.TypeOf<Run>());
                Assert.That(elements[0].InnerText, Is.EqualTo("Lorem"));
                Assert.That(elements[1].ChildElements, Has.All.TypeOf<Run>());
                Assert.That(elements[1].InnerText, Is.EqualTo("Placeholder"));
                Assert.That(elements[2].ChildElements, Has.All.TypeOf<Run>());
                Assert.That(elements[2].InnerText, Is.EqualTo("Ipsum"));
            });
            Assert.That(elements[1].LastChild?.HasChild<LastRenderedPageBreak>(), Is.False);
        }

        [TestCase("rtl", ExpectedResult = true)]
        [TestCase("ltr", ExpectedResult = false)]
        [TestCase("", ExpectedResult = null)]
        public bool? WithRtl_ReturnsBidi(string dir)
        {
            var elements = converter.Parse($@"<div dir='{dir}'>Lorem</div>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            var bidi = elements[0].GetFirstChild<ParagraphProperties>()?.BiDi;
            return bidi?.Val?.Value;
        }

        [TestCase("1.5", "auto", "360", Description = "Unitless")]
        [TestCase("150%", "auto", "360")]
        [TestCase("100%", "auto", "240")]
        [TestCase("25px", "exact", "375")]
        [TestCase("3em", "exact", "720")]
        [TestCase("normal", "auto", "240", Description = "Depend on the user agent")]
        public void WithLineHeight_ReturnsSpacingBetweenLines(string lineHeight, string expectedRule, string expectedSpace)
        {
            var elements = converter.Parse($@"<div style='line-height: {lineHeight}'>
                Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer accumsan placerat sem in consequat. Quisque viverra ex elit, et volutpat libero finibus eget. Vivamus placerat velit ex, quis rutrum justo molestie eget.
            </div>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            var spaces = elements[0].GetFirstChild<ParagraphProperties>()?.SpacingBetweenLines;
            Assert.That(spaces?.LineRule?.InnerText, Is.EqualTo(expectedRule));
            Assert.That(spaces?.Line?.Value, Is.EqualTo(expectedSpace));
        }

        [Test(Description = "Block endings with line break, should ignore it #158")]
        public void WithEndingLineBreak_ReturnsIgnoredBreak()
        {
            var elements = converter.Parse("line1<div>line2<br><div>line3<br></div></div>");
            Assert.That(elements, Has.Count.EqualTo(3));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            Assert.That(elements.Any(p => p.LastChild?.LastChild is Break), Is.False);
        }

        [Test(Description = "Block endings with 2 line breaks, should keep only one")]
        public void WithEndingDoubleLineBreak_ReturnsOneBreak()
        {
            var elements = converter.Parse("line1<div>line2<br><br><div>line3<br></div></div>");
            Assert.That(elements, Has.Count.EqualTo(3));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            Assert.That(elements.ElementAt(1).LastChild?.LastChild, Is.TypeOf<Break>());
        }

        [Test(Description = "Block containing only 1 line break, should display empty run")]
        public void WithOnlyLineBreak_ReturnsEmptyRun()
        {
            var elements = converter.Parse("<div><br></div>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            var lastRun = elements.First().LastChild;
            Assert.That(lastRun, Is.Not.Null);
            Assert.Multiple(() => {
                Assert.That(lastRun.LastChild, Is.Not.TypeOf<Break>());
                Assert.That(lastRun.LastChild, Is.TypeOf<Text>());
            });
            Assert.That(((Text)lastRun.LastChild).Text, Is.Empty);
        }

        [Test(Description = "Border defined on container should render its content with one bordered frame #168")] 
        public async Task WithBorders_MultipleParagraphs_ReturnsAsOneFramedBlock()
        {
            await converter.ParseBody(@"<div style='margin-top: 20px; border: 1px dashed rgba(0, 0, 0, 0.4); padding: 6px 8px; font-size: 14px;'>
              <div>
                <p>Header placeholder:</p>
                <ol>
                    <li>Item 1</li>
                    <li>Item 2</li>
                </ol>
                <p style='text-indent: 4.5em'>Footer Placeholder</p>
              </div>
            </div>");
            AssertThatOpenXmlDocumentIsValid();

            var paragraphs = mainPart.Document.Body!.Elements<Paragraph>();
            Assert.That(paragraphs, Is.Empty, "Assert that all the paragraphs stand inside the framed table");

            var framedTable = mainPart.Document.Body!.Elements<Table>().FirstOrDefault();
            Assert.That(framedTable, Is.Not.Null);

            var borders = framedTable.GetFirstChild<TableProperties>()?.TableBorders;
            Assert.That(borders, Is.Not.Null, "Assert that border is applied on table scope");
            Assert.That(borders.Elements<BorderType>()!
                .Select(b => b.Val?.Value),
                Has.All.EqualTo(BorderValues.Dashed));

            var cell = framedTable.GetFirstChild<TableRow>()?.GetFirstChild<TableCell>();
            Assert.That(cell, Is.Not.Null);
            paragraphs = cell.Elements<Paragraph>();
            Assert.That(paragraphs, Is.Not.Empty);

            Assert.That(paragraphs.Last().ParagraphProperties?.Indentation?.FirstLine?.Value, Is.EqualTo("1080"),
                "Assert that paragraph with text-indent is preserved");
            Assert.That(paragraphs.Last().ParagraphProperties?.Indentation?.Right, Is.Null,
                "Assert that paragraph with right indentation is preserved");
        }

        [Test(Description = "Background color defined on container should render its content with one bordered frame")] 
        public async Task WithBgcolor_MultipleParagraphs_ReturnsAsOneFramedBlock()
        {
            await converter.ParseBody(@"<article style='background: orange'>
                <header>Header placeholder</header>
                <p>Body Placeholder</p>
            </article>");
            AssertThatOpenXmlDocumentIsValid();

            var paragraphs = mainPart.Document.Body!.Elements<Paragraph>();
            Assert.That(paragraphs, Is.Empty, "Assert that all the paragraphs stand inside the framed table");

            var framedTable = mainPart.Document.Body!.Elements<Table>().FirstOrDefault();
            Assert.That(framedTable, Is.Not.Null);

            var shading = framedTable.GetFirstChild<TableRow>()?.GetFirstChild<TableCell>()?.TableCellProperties?.Shading;
            Assert.That(shading, Is.Not.Null, "Assert that background-color is applied on table scope");
            Assert.That(shading.Fill?.Value, Is.EqualTo("FFA500"));
        }
    }
}