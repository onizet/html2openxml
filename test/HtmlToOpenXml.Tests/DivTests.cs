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
            var elements = converter.Parse(@"<div style='text-indent:1em;border:1px dotted red;text-align:center'>Lorem</div>");
            Assert.That(elements, Has.Count.EqualTo(1));
            Assert.That(elements, Has.All.TypeOf<Paragraph>());
            var p = (Paragraph) elements[0];
            Assert.Multiple(() =>
            {
                Assert.That(p.ParagraphProperties?.Indentation?.FirstLine?.HasValue, Is.True);
                Assert.That(p.ParagraphProperties?.ParagraphBorders, Is.Not.Null);
                Assert.That(p.ParagraphProperties?.Justification?.Val?.Value, Is.EqualTo(JustificationValues.Center));
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
            Assert.That(elements[1].LastChild?.HasChild<Break>(), Is.True);
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
    }
}